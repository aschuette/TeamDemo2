<!--#include file ="checkusr.asp"-->
<!--#include file ="vb.asp"-->
<!--#include file ="db.asp"-->
<%

set rs = Server.CreateObject("ADODB.RecordSet")

id = request("id")

task = request("task")

if task = "" then
  drawSessionZone = true
  changeSessionZone = "document.location='p_sets.asp'"
  PrintShell
  response.end
elseif task = "newmyset" or task = "newcommon" then
  sql = "INSERT INTO student_sets (is_common,zone_name,owner,set_name,set_size,size_calc_date,created_by,created_dt,updated_by,updated_dt) "
  sql = sql & " VALUES ("
  if task = "newcommon" then
    sql = sql & "'Y',"
  else
    sql = sql & "'N',"
  end if
  sql = sql & checkstring(session("zone"),50) & "," & checkstring(session("user_id"),50) & "," & checkstring(request("name"),50) & ",'0'," & checkstring(now,50) & ","
  sql = sql & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ", "
  sql = sql & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ")"
  conn.execute(sql)
  randomize
  r = int(rnd()*999999)
  response.redirect("p_sets.asp?r="&r)
elseif task = "search" then
  SearchList
elseif task = "menu" then
  SetMenu
elseif task = "listreportshell" then
  ListReportShell
elseif task = "listreportfields" then
  ListReportFields
elseif task = "viewlistreport" then
  ViewListReport
elseif task = "billingshell" then
  BillingShell
elseif task = "viewbilling" then
  ViewBilling
elseif task = "savebilling" then
  SaveBilling
elseif task = "viewadvisors" then
  ViewAdvisors
elseif task = "saveadvisors" then
  SaveAdvisors
elseif task = "viewregstatus" then
  ViewRegStatus
elseif task = "saveregstatus" then
  SaveRegStatus
elseif task = "requirementsshell" then
  RequirementsShell
elseif task = "viewrequirements" then
  ViewRequirements
elseif task = "saverequirements" then
  SaveRequirements
elseif task = "requirementsreport" then
  RequirementsReport
elseif task = "viewcorrespondence" then
  ViewCorrespondence
elseif task = "savecorrespondence" then
  SaveCorrespondence
else
  ShowList
end if




''***************************************************
''**                MAIN FUNCTIONS                 **
''***************************************************




sub ShowList
  response.write("<html>")
  response.write("<head>")
  response.write("<title>Search</title>")
  response.write("<style>")
  response.write("BODY body { scrollbar-arrow-color: #990000; scrollbar-3dlight-color: #666666; scrollbar-highlight-color: #DDDDDD; scrollbar-face-color: #DDDDDD; scrollbar-shadow-color: #DDDDDD; scrollbar-track-color: #EEEEEE; scrollbar-darkshadow-color: #666666}")
  response.write("A:link#alink2 {text-decoration: none; color: #000000}")
  response.write("A:visited#alink2 {text-decoration: none; color: #000000}")
  response.write("A:active#alink2 {text-decoration: none; color: #000000}")
  response.write("A:hover#alink2 {text-decoration: underline; color: #990000}")
  response.write("</style>")
  response.write("</head>")
  response.write("<body bgcolor=F9F9F9 text=000000 link=000000 vlink=000000 alink=000000 topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>")
  response.write("<center>")
  response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% height=100% ><tr><td valign=middle align=center>")
  sql = "select * from student_sets WHERE (is_common='Y' OR owner=" & checkstring(session("user_id"),50) & ")"
  if session("zone") <> "" then sql = sql & " AND zone_name = " & checkstring(session("zone"),50)
  sql = sql & " order by is_common, set_name"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write("<font face='arial,helvetica' size=-1>No student sets are available.</font><br>")
    rs.close
  else
    response.write("<script>function SelectSet(which) {parent.document.location='p_sets.asp?task=menu&id='+which;}</script>")
    response.write("<script>function GoSet(which) {parent.document.location='p_set_def.asp?id='+which;}</script>")
    response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td align=center bgcolor=999999>")
    response.write("<table border=0 cellspacing=1 cellpadding=1>")
    response.write("<tr>")
    response.write("<td align=center bgcolor=EEEEEE><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<b>Type</b>&nbsp;&nbsp;</font></nobr></td>")
    response.write("<td align=center bgcolor=EEEEEE><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<b>Name</b>&nbsp;&nbsp;</font></nobr></td>")
    response.write("<td align=center bgcolor=EEEEEE><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<b>Size</b>&nbsp;&nbsp;</font></nobr></td>")
    response.write("<td align=center bgcolor=EEEEEE><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<b>Updated</b>&nbsp;&nbsp;</font></nobr></td>")
    response.write("<td align=center bgcolor=EEEEEE><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<b>Details</b>&nbsp;&nbsp;</font></nobr></td>")
    response.write("</tr>")
    while not rs.eof
      response.write("<tr>")
      if rs("is_common") = "Y" then
        MySet = "Common"
      else
        MySet = "MySet"
      end if
      response.write("<td bgcolor=FFFFFF align=center><font face='arial,helvetica' size=-1>&nbsp;" & MySet & "&nbsp;</font></td>")
      response.write("<td bgcolor=FFFFFF align=center><font face='arial,helvetica' size=-1>&nbsp;<a href=""JavaScript:SelectSet('" & rs("student_set_id") & "');""><font color=990000>" & rs("set_name") & "</font></a>&nbsp;</font></td>")
      response.write("<td bgcolor=FFFFFF align=center><font face='arial,helvetica' size=-1>" & rs("set_size") & "</font></td>")
      updated_dt = rs("updated_dt")
      pos = instr(updated_dt," ")
      if pos > 0 then updated_dt = left(updated_dt,pos-1)
      response.write("<td bgcolor=FFFFFF align=center><font face='arial,helvetica' size=-1>" & updated_dt & "</font></td>")
      response.write("<td bgcolor=FFFFFF align=center><font face='arial,helvetica' size=-1><a href=""JavaScript:GoSet('" & rs("student_set_id") & "');""><img src=""images/info.gif"" border=0 width=14 height=14></a></font></td>")
      rs.movenext
    wend
    rs.close
    response.write("</td></tr></table>")
    response.write("</td></tr></table>")
  end if
  response.write("</td></tr></table>")
  response.write("</center>")
  response.write("</body>")
  response.write("</html>")
end sub





sub PrintShell
session("pagetab") = "People"
session("history") = "Menu|p_menu.asp|Student Sets"
PageHeader
%>
<script>
function NewMySet() {
  x = prompt('Enter a name for the student set.','Untitled Set');
  if (x) {
    document.location='p_sets.asp?task=newmyset&name='+x;
  }
}
function NewCommon() {
  x = prompt('Enter a name for the student set.','Untitled Set');
  if (x) {
    document.location='p_sets.asp?task=newcommon&name='+x;
  }
}
</script>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left valign=middle height=100%>
<iframe name="list" src="p_sets.asp?task=view" scrolling=auto width=100% height=100% frameborder=0></iframe>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:110px; background-color:#CCCCCC" type=button onClick="NewMySet();" name="bmyset" value=" New MySet ">
<img src="images/spacer.gif" border=0 width=10 height=1>
<% if isRole("U") then %>
<input style="width:110px; background-color:#CCCCCC" type=button onClick="NewCommon();" name="bcommon" value=" New Common ">
<img src="images/spacer.gif" border=0 width=10 height=1>
<% end if %>
<input style="width:110px; background-color:#CCCCCC" type=button onclick="GoPage('p_menu.asp');" name="bcancel" value=" Back ">
</nobr></td>
<%
PageFooter
end sub





sub SearchList
  response.write("<html>")
  response.write("<head>")
  response.write("<title>Search</title>")
  response.write("<style>")
  response.write("BODY body { scrollbar-arrow-color: #990000; scrollbar-3dlight-color: #666666; scrollbar-highlight-color: #DDDDDD; scrollbar-face-color: #DDDDDD; scrollbar-shadow-color: #DDDDDD; scrollbar-track-color: #EEEEEE; scrollbar-darkshadow-color: #666666}")
  response.write("A:link#alink2 {text-decoration: none; color: #000000}")
  response.write("A:visited#alink2 {text-decoration: none; color: #000000}")
  response.write("A:active#alink2 {text-decoration: none; color: #000000}")
  response.write("A:hover#alink2 {text-decoration: underline; color: #990000}")
  response.write("</style>")
  response.write("</head>")
  response.write("<body bgcolor=FFFFFF text=000000 link=000000 vlink=000000 alink=000000 topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>")
  sql = "select * from student_sets WHERE (is_common='Y' OR owner=" & checkstring(session("user_id"),50) & ")"
  if session("zone") <> "" then sql = sql & " AND zone_name = " & checkstring(session("zone"),50)
  sql = sql & " order by is_common, set_name"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write("<table border=0 cellspacing=0 cellpadding=5><tr><td><font face='arial' size=-1 color=000000>")
    response.write("No student sets are available.>")
    response.write("</font></td></tr></table>")
    rs.close
  else
    n = rs.RecordCount
    if n = 1 then
      s = "1 student set found"
    else
      s = n & " student sets found"
    end if
    response.write("<script>parent.document.getElementById(""searchcount"").innerHTML = '" & s & "&nbsp;';</script>")
    response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% ><tr><td bgcolor=CCCCCC>")
    response.write("<table border=0 cellspacing=1 cellpadding=0 width=100% >")
    response.write("<tr bgcolor=EEEEEE>")
    response.write("<td align=left><nobr><font face='arial,helvetica' size=-1>&nbsp;Type&nbsp;</font></nobr></td>")
    response.write("<td align=left><nobr><font face='arial,helvetica' size=-1>&nbsp;Name&nbsp;</font></nobr></td>")
    response.write("<td align=left><nobr><font face='arial,helvetica' size=-1>&nbsp;Size&nbsp;</font></nobr></td>")
    response.write("</tr>")
    while not rs.eof
      response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" onClick=""parent.addSet('" & rs("student_set_id") & "');"">")
      if rs("is_common") = "Y" then
        MySet = "Common"
      else
        MySet = "MySet"
      end if
      response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & MySet & "&nbsp;</font></td>")
      response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("set_name") & "&nbsp;</font></td>")
      response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("set_size") & "&nbsp;</font></td>")
      response.write("</tr>")
      rs.movenext
    wend
    rs.close
    response.write("</td></tr></table>")
    response.write("</td></tr></table>")
  end if
  response.write("</body>")
  response.write("</html>")
end sub


sub SetMenu
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes"
  PageHeader
  sql = "select set_name, set_query from student_sets where student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_name = rs("set_name")
    set_query = rs("set_query")
  end if
  rs.close
  if isnull(set_name) then set_name = ""
  if set_name = "" then  set_name = "Untitled Set"
  if isnull(set_query) then set_query = ""
  if set_query = "" then set_query = "select 0 student_instance_id where 1=0"
  sql = "select count(*) n from (" & set_query & ") t"
  rs.open sql,conn,1,1
  if not rs.eof then
    set_size = rs("n")
  end if
  rs.close
  if isnull(set_size) then set_size = ""
  if set_size = "" then set_size = "0"
  if set_size <> "" then
    sql = "update student_sets set set_size = " & checkstring(set_size,50) & ", size_calc_date = GetDate() where student_set_id = " & checkstring(id,50)
    conn.execute(sql)
  end if
  if set_size = "1" then
    set_size = set_size & "&nbsp;student"
  else
    set_size = set_size & "&nbsp;students"
  end if
  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td align=left bgcolor=CCCCCC>")
  response.write("<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>")
  response.write("<font face='arial,helvetica'>")
  response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% ><tr>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' color=666666><b>" & set_name & "</b></font></td>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;" & set_size & "</font></td>")
  response.write("</tr></table>")
  response.write("<img src=""images/spacer.gif"" border=0 width=300 height=15><br>")
  response.write("<center>")
%>

<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table width=100% border=0 cellspacing=1 cellpadding=5><tr><td bgcolor=DDDDDD align=center valign=middle onMouseOver="this.bgColor='#FFFFCC';" onMouseOut="this.bgColor='#DDDDDD';" onClick="document.location='p_sets.asp?task=listreportshell&id=<%=id%>';">
<font face='arial,helvetica'><b>Student Listing</b><br></font>
</td></tr></table>
</td></tr></table>

<% if isRole("U") then %>

<img src="images/spacer.gif" border=0 width=1 height=10><br>

<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table width=100% border=0 cellspacing=1 cellpadding=5><tr><td bgcolor=DDDDDD align=center valign=middle onMouseOver="this.bgColor='#FFFFCC';" onMouseOut="this.bgColor='#DDDDDD';" onClick="document.location='p_sets.asp?task=billingshell&id=<%=id%>';">
<font face='arial,helvetica'><b>Term Billing</b><br></font>
</td></tr></table>
</td></tr></table>

<img src="images/spacer.gif" border=0 width=1 height=10><br>

<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table width=100% border=0 cellspacing=1 cellpadding=5><tr><td bgcolor=DDDDDD align=center valign=middle onMouseOver="this.bgColor='#FFFFCC';" onMouseOut="this.bgColor='#DDDDDD';" onClick="document.location='p_sets.asp?task=viewadvisors&id=<%=id%>';">
<font face='arial,helvetica'><b>Advisor Assignment</b><br></font>
</td></tr></table>
</td></tr></table>

<img src="images/spacer.gif" border=0 width=1 height=10><br>

<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table width=100% border=0 cellspacing=1 cellpadding=5><tr><td bgcolor=DDDDDD align=center valign=middle onMouseOver="this.bgColor='#FFFFCC';" onMouseOut="this.bgColor='#DDDDDD';" onClick="document.location='p_sets.asp?task=viewregstatus&id=<%=id%>';">
<font face='arial,helvetica'><b>Registration Status</b><br></font>
</td></tr></table>
</td></tr></table>

<img src="images/spacer.gif" border=0 width=1 height=10><br>

<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table width=100% border=0 cellspacing=1 cellpadding=5><tr><td bgcolor=DDDDDD align=center valign=middle onMouseOver="this.bgColor='#FFFFCC';" onMouseOut="this.bgColor='#DDDDDD';" onClick="document.location='p_sets.asp?task=requirementsshell&id=<%=id%>';">
<font face='arial,helvetica'><b>Requirements</b><br></font>
</td></tr></table>
</td></tr></table>

<% end if %>

<img src="images/spacer.gif" border=0 width=1 height=10><br>

<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>
<table width=100% border=0 cellspacing=1 cellpadding=5><tr><td bgcolor=DDDDDD align=center valign=middle onMouseOver="this.bgColor='#FFFFCC';" onMouseOut="this.bgColor='#DDDDDD';" onClick="document.location='p_sets.asp?task=viewcorrespondence&id=<%=id%>';">
<font face='arial,helvetica'><b>Correspondences</b><br></font>
</td></tr></table>
</td></tr></table>

<%
  response.write("</font>")
  response.write("</center>")
  response.write("</td></tr></table>")
  response.write("</td></tr></table>")
  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")
  response.write("<font face='arial,helvetica' size=-1><nobr>&nbsp;<b>Note:</b> Batch processes can unexpectedly change who is in a person set.&nbsp;</nobr></font><br>")
  PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:80px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp');" name="bcancel" value=" Back ">
</nobr></td>
<%
  PageFooter
end sub



sub SaveComplete(message)
  PageHeader
%>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=center valign=middle height=100%>
<%
  sql = "select set_name from student_sets where student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_name = rs("set_name")
  end if
  rs.close
  if isnull(set_name) then set_name = ""
  if set_name = "" then  set_name = "Untitled Set"
  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign=middle>")
  response.write("<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>")
  response.write("<font face='arial,helvetica'>")
  response.write("<font color=666666><b>" & set_name & "</b></font><br>")
  response.write("<img src=""images/spacer.gif"" border=0 width=300 height=15><br>")
  response.write("<font size=-1>")
  response.write("<center>")
  response.write(message & "<br>")
  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=5><br>")
  response.write("Click the Continue button below for other batch options.<br>")
  response.write("</center>")
  response.write("</font>")
  response.write("</td></tr></table>")
  response.write("</td></tr></table>")
%>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:150px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp?task=menu&id=<%=id%>');" name="bcancel" value=" Continue ">
</nobr></td>
<%
PageFooter
end sub





''***************************************************
''**                 STUDENT LISTS                 **
''***************************************************




sub ListReportShell
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Student Listing"
  PageHeader
  if request("task2") = "cols" then
    cols = "Z"
    for i = 0 to 15
      cols = cols & request("F"&i)
    next
  else
    cols = request("cols")
  end if
%>
<script>
function DoReportPrint() {
  document.details.focus();
  document.details.print();
}
function GoReportExcel() {
  window.open("p_sets.asp?task=viewlistreport&type=excel&id=<%=id%>&cols=<%=cols%>","StudentSetReport","resizable=yes,width=680,height=400,scrollbars,menubar,toolbar,location,status");
}
function GoReportText() {
  //window.open("p_sets.asp?task=viewlistreport&type=text&id=<%=id%>","StudentSetReport","resizable=yes,width=680,height=400,scrollbars,menubar,toolbar,location,status");
  document.location = "p_sets.asp?task=viewlistreport&type=text&id=<%=id%>&cols=<%=cols%>","StudentSetReport","resizable=yes,width=680,height=400,scrollbars,menubar,toolbar,location,status";
}
function GoReportQuery() {
  window.open("p_sets.asp?task=viewlistreport&type=query&id=<%=id%>&cols=<%=cols%>","StudentSetQuery","resizable=yes,width=400,height=200,scrollbars,menubar,status");
}
function GoReportFields() {
  document.location = "p_sets.asp?task=listreportfields&id=<%=id%>&cols=<%=cols%>";
}
</script>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left valign=middle height=100%>
<iframe name="details" src="p_sets.asp?task=viewlistreport&id=<%=id%>&cols=<%=cols%>" scrolling=auto width=100% height=100% frameborder=0></iframe>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:55px; background-color:#CCCCCC" type=button onClick="GoPage('p_set_def.asp?id=<%=id%>');" name="bedit" value=" Edit ">
<img src="images/spacer.gif" border=0 width=5 height=1>
<input style="width:55px; background-color:#CCCCCC" type=button onClick="DoReportPrint();" name="bprint" value=" Print ">
<img src="images/spacer.gif" border=0 width=5 height=1>
<input style="width:55px; background-color:#CCCCCC" type=button onClick="GoReportExcel();" name="bexcel" value=" Excel ">
<img src="images/spacer.gif" border=0 width=5 height=1>
<input style="width:55px; background-color:#CCCCCC" type=button onClick="GoReportText();" name="btext" value=" Text ">
<img src="images/spacer.gif" border=0 width=5 height=1>
<input style="width:55px; background-color:#CCCCCC" type=button onClick="GoReportQuery();" name="bquery" value="Query">
<img src="images/spacer.gif" border=0 width=5 height=1>
<input style="width:55px; background-color:#CCCCCC" type=button onClick="GoReportFields();" name="bfields" value="Fields">
<img src="images/spacer.gif" border=0 width=5 height=1>
<input style="width:55px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp?task=menu&id=<%=id%>');" name="bcancel" value=" Back ">
</nobr></td>
<%
PageFooter
end sub


sub ListReportFields
  cols = request("cols")
  if cols = "" then cols = "ZTHL"
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Student Listing|p_sets.asp?task=listreportshell&id=" & id & "&cols=" & cols & "|Select Fields"
  PageHeader
%>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left valign=middle height=100%>
<center>
<table border=0 cellspacing=0 cellpadding=0><tr><td><form name=colform method=GET action="p_sets.asp" target="_parent"><input type=hidden name="task" value="listreportshell"><input type=hidden name=id value="<%=id%>"><input type=hidden name="task2" value="cols"></td><td align=center bgcolor=CCCCCC>
<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE><font face='arial,helvetica'>
<font color=666666><b>Select Report Fields</b></font><br>
<center>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr><td valign=middle align=center><input type=checkbox name="F0" value="F" <% if instr(cols,"F") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Full Name</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F1" value="T" <% if instr(cols,"T") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Program Track</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F2" value="H" <% if instr(cols,"H") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Harvard ID</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F3" value="N" <% if instr(cols,"N") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>SSN</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F4" value="B" <% if instr(cols,"B") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>DOB</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F5" value="M" <% if instr(cols,"M") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Matric Dt</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F6" value="G" <% if instr(cols,"G") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Exp Grad Dt</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F7" value="Y" <% if instr(cols,"Y") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>YOS</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F8" value="X" <% if instr(cols,"X") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Gender</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F9" value="E" <% if instr(cols,"E") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Ethnicity (Old)</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F10" value="C" <% if instr(cols,"C") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Citizenship</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F11" value="S" <% if instr(cols,"S") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Society</font></td></tr>
<tr><td valign=middle align=center><input type=checkbox name="F12" value="L" <% if instr(cols,"L") > 0 then %>checked<% end if %>></td><td align=left valign=middle><font face='arial,helvetica' size=-1>Email</font></td></tr>
<tr><td valign=center align=center><input type=checkbox name="F13" value="I" <% if instr(cols,"I") > 0 then %>checked<% end if %>></td><td align=left valign=center><font face='arial,helvetica' size=-1>Ethnicity (Ipeds)</font></td></tr>
<tr><td valign=center align=center><input type=checkbox name="F14" value="D" <% if instr(cols,"D") > 0 then %>checked<% end if %>></td><td align=left valign=center><font face='arial,helvetica' size=-1>DentPin</font></td></tr>
<tr><td valign=center align=center><input type=checkbox name="F15" value="A" <% if instr(cols,"A") > 0 then %>checked<% end if %>></td><td align=left valign=center><font face='arial,helvetica' size=-1>AAMC ID</font></td></tr>
</table>
</center>
</td></tr></table>
</td><td></form></td></tr></table>
</center>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:55px; background-color:#CCCCCC" type=button onClick="document.colform.submit()" name="bsave" value=" Save ">
<img src="images/spacer.gif" border=0 width=5 height=1>
<input style="width:55px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp?task=listreportshell&id=<%=id%>&cols=<%=cols%>');" name="bcancel" value=" Back ">
</nobr></td>
<%
PageFooter
end sub



sub ViewListReport
  cols = request("cols")
  if cols = "" then cols = "ZTHL"
  sql = "select set_name, set_size, set_query from student_sets where student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_name = rs("set_name")
    set_query = rs("set_query")
    set_size = rs("set_size")
  end if
  rs.close
  if isnull(set_name) then set_name = ""
  if set_name = "" then  set_name = "Untitled Set"
  if isnull(set_size) then set_size = ""
  if set_size = "" then set_size = "0"
  if set_size = "1" then
    set_size = set_size & "&nbsp;student"
  else
    set_size = set_size & "&nbsp;students"
  end if
  if isnull(set_query) then set_query = ""
  if set_query = "" then set_query = "select 0 student_instance_id where 1=0"
  sql = "SELECT p.full_name, p.last_name, p.first_name, p.middle_name, p.harvard_id, i.student_instance_id, g.program_name, t.track_name, en.hispanic_flag, en.amind_flag, en.asian_flag, en.black_flag, en.pacif_flag, en.white_flag, MAL.amcas_id "
  sql = sql & ", p.ssn, p.date_of_birth, p.email_address, p.dentpin, i.matric_date, i.expected_grad_date, y.year_of_study_name, p.gender, p.citizenship, i.society "
  sql = sql & ", (select top 1 e2.ethnicity_code from person_ethnicity as e1, ethnicity as e2 where e1.primary_flag = 'Y' and e1.person_id = p.person_id and e1.ethnicity_id = e2.ethnicity_id) ethnicity "
  sql = sql & " FROM person p LEFT OUTER JOIN person_ethnicity_new en  ON P.person_id = en.person_id LEFT OUTER JOIN Madris_Amcas_Lookup MAL ON P.person_id = MAL.person_id,  student_instance i left outer join year_of_study y on i.year_of_study_id = y.year_of_study_id,  programs g, program_track t "
  sql = sql & " WHERE p.person_id = i.person_id and g.program_id = i.program_id and t.program_track_id = i.program_track_id "
  sql = sql & " and i.student_instance_id in (" & set_query & ")"
  if instr(cols,"F") > 0 then
    sql = sql & " order by p.full_name "
  else
    sql = sql & " order by p.last_name, p.first_name, p.middle_name "
  end if
  if request("type") = "text" then
    response.ContentType = "text/xml"
    response.AddHeader "Content-Disposition", "attachment; filename=""" & replace(set_name," ","_") & ".txt"""
    rs.open sql,conn,1,1
    if rs.eof then
      response.write("There are no students in this set.")
    else
      while not rs.eof
        huid = rs("harvard_id")
        if isnull(huid) then huid = ""
        if instr(cols,"F") > 0 then
          response.Write(rs("full_name"))
        else
          response.write(rs("last_name"))
          response.write(vbTab & rs("first_name"))
          response.write(vbTab & rs("middle_name"))
        end if
        if instr(cols,"T") > 0 then response.write(vbTab & rs("program_name") & " (" & rs("track_name") & ")")
        if instr(cols,"H") > 0 then response.write(vbTab & huid)
        if instr(cols,"N") > 0 then response.write(vbTab & rs("ssn"))
        if instr(cols,"B") > 0 then response.write(vbTab & rs("date_of_birth"))
        if instr(cols,"M") > 0 then response.write(vbTab & rs("matric_date"))
        if instr(cols,"G") > 0 then response.write(vbTab & rs("expected_grad_date"))
        if instr(cols,"Y") > 0 then response.write(vbTab & rs("year_of_study_name"))
        if instr(cols,"X") > 0 then response.write(vbTab & rs("gender"))
        if instr(cols,"E") > 0 then response.write(vbTab & rs("ethnicity"))
        if instr(cols,"C") > 0 then response.write(vbTab & rs("citizenship"))
        if instr(cols,"S") > 0 then response.write(vbTab & rs("society"))
        if instr(cols,"L") > 0 then response.write(vbTab & rs("email_address"))
		if instr(cols,"I") > 0 then
			response.write(vbTab & rs("hispanic_flag"))
			response.write(vbTab & rs("amind_flag"))
			response.write(vbTab & rs("asian_flag"))
			response.write(vbTab & rs("black_flag"))
			response.write(vbTab & rs("pacif_flag"))
			response.write(vbTab & rs("white_flag"))
		end if
         if instr(cols,"D") > 0 then response.write(vbTab & rs("dentpin"))
			if instr(cols,"A") > 0 then response.write(vbTab & rs("amcas_id"))
        response.write(vbCrLf)
        rs.movenext
      wend
      rs.close
    end if
  elseif request("type") = "excel" then
    response.ContentType = "application/vnd.ms-excel"
    response.AddHeader "Content-Disposition", "filename=""" & replace(set_name," ","_") & ".xls"""
    rs.open sql,conn,1,1
    if rs.eof then
      response.write("There are no students in this set.")
    else
      response.write("<table border=1>")
      response.write("<tr>")
      if instr(cols,"F") > 0 then
        response.write("<td><b>Name</b></td>")
      else
        response.write("<td><b>Last</b></td>")
        response.write("<td><b>First</b></td>")
        response.write("<td><b>Middle</b></td>")
      end if
      if instr(cols,"T") > 0 then response.write("<td><b>Program (Track)</b></td>")
      if instr(cols,"H") > 0 then response.write("<td><b>Harvard ID</b></td>")
      if instr(cols,"N") > 0 then response.write("<td><b>SSN</b></td>")
      if instr(cols,"B") > 0 then response.write("<td><b>DOB</b></td>")
      if instr(cols,"M") > 0 then response.write("<td><b>Matric Dt</b></td>")
      if instr(cols,"G") > 0 then response.write("<td><b>Exp Grad Dt</b></td>")
      if instr(cols,"Y") > 0 then response.write("<td><b>YOS</b></td>")
      if instr(cols,"X") > 0 then response.write("<td><b>Gender</b></td>")
      if instr(cols,"E") > 0 then response.write("<td><b>Ethnicity</b></td>")
      if instr(cols,"C") > 0 then response.write("<td><b>Citizenship</b></td>")
      if instr(cols,"S") > 0 then response.write("<td><b>Society</b></td>")
      if instr(cols,"L") > 0 then response.write("<td><b>Email</b></td>")
	  if instr(cols,"I") > 0 then
	  		response.write("<td><b>Hispanic</b></td>")
			response.write("<td><b>American Indian or Alaskan Native</b></td>")
			response.write("<td><b>Asian</b></td>")
			response.write("<td><b>Black or African American</b></td>")
			response.write("<td><b>Native Hawaiian or Pacific Islander</b></td>")
			response.write("<td><b>White</b></td>")
		end if
        if instr(cols,"D") > 0 then response.write("<td><b>DentPin</b></td>")
		  if instr(cols,"A") > 0 then response.write("<td><b>AAMC ID</b></td>")
      response.write("</tr>")
      while not rs.eof
        huid = rs("harvard_id")
        if isnull(huid) then huid = ""
        response.write("<tr>")
        if instr(cols,"F") > 0 then
          response.write("<td>" & rs("full_name") & "</td>")
        else
          response.write("<td>" & rs("last_name") & "</td>")
          response.write("<td>" & rs("first_name") & "</td>")
          response.write("<td>" & rs("middle_name") & "</td>")
        end if
        if instr(cols,"T") > 0 then response.write("<td>" & rs("program_name") & " (" & rs("track_name") & ")")
        if instr(cols,"H") > 0 then response.write("<td>" & huid & "</td>")
        if instr(cols,"N") > 0 then response.write("<td>" & rs("ssn") & "</td>")
        if instr(cols,"B") > 0 then response.write("<td>" & rs("date_of_birth") & "</td>")
        if instr(cols,"M") > 0 then response.write("<td>" & rs("matric_date") & "</td>")
        if instr(cols,"G") > 0 then response.write("<td>" & rs("expected_grad_date") & "</td>")
        if instr(cols,"Y") > 0 then response.write("<td>" & rs("year_of_study_name") & "</td>")
        if instr(cols,"X") > 0 then response.write("<td>" & rs("gender") & "</td>")
        if instr(cols,"E") > 0 then response.write("<td>" & rs("ethnicity") & "</td>")
        if instr(cols,"C") > 0 then response.write("<td>" & rs("citizenship") & "</td>")
        if instr(cols,"S") > 0 then response.write("<td>" & rs("society") & "</td>")
        if instr(cols,"L") > 0 then response.write("<td>" & rs("email_address") & "</td>")
		 if instr(cols,"I") > 0 then
	  		response.write("<td>"  & rs("hispanic_flag") & "</td>")
			response.write("<td>"  & rs("amind_flag") & "</td>")
			response.write("<td>"  & rs("asian_flag") & "</td>")
			response.write("<td>"  & rs("black_flag") & "</td>")
			response.write("<td>"  & rs("pacif_flag") & "</td>")
			response.write("<td>"  & rs("white_flag") & "</td>")
		end if
         if instr(cols,"D") > 0 then response.write("<td>" & rs("dentpin") & "</td>")
			if instr(cols,"A") > 0 then response.write("<td>" & rs("amcas_id") & "</td>")
        response.write("</tr>")
        rs.movenext
      wend
      response.write("</table>")
      rs.close
    end if
  elseif request("type") = "query" then
%>
<html>
<META HTTP-EQUIV="Expires" Content="0">
<META HTTP-EQUIV="Pragma" Content="no-cache">
<META HTTP-EQUIV="Cache-Control" Content="Private">
<head>
<title>MADRIS</title>
</head>
<body bgcolor="EEEEEE" text="000000" link="000000" vlink="000000" alink="000000">
<center>
<table border=0 cellspacing=0 cellpadding=0><tr><td><form></td><td align=center>
<font face='arial,helvetica' color=666666><b><%=set_name%></b></font><br>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<textarea rows=6 cols=40><%=set_query%></textarea><br>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<input type=button value=" Close " onClick="window.close();"><br>
</td><td></form></td></tr></table>
</center>
</body>
</html>
<%
  else
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
function DoSave() {
  document.dataform.task.value = "savebilling";
  document.dataform.submit();
}
</script>
</head>
<body bgcolor="F9F9F9" text="000000" link="000000" vlink="000000" alink="000000">
<%
    response.write("<font size=-1>")
    response.write("<center>")
    response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td valign=middle>")
    response.write("<font face='arial,helvetica'>")
    response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% ><tr>")
    response.write("<td align=left valign=middle><font face='arial,helvetica' color=666666><b>" & set_name & "</b></font></td>")
    response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;" & set_size & "</font></td>")
    response.write("</tr></table>")
    response.write("<img src=""images/spacer.gif"" border=0 width=300 height=15><br>")
    response.write("<center>")
    rs.open sql,conn,1,1
    if rs.eof then
      response.write("<font size=-1><i>There are no students in this set.</i></font><br>")
    else
      response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=999999>")
      response.write("<table border=0 cellspacing=1 cellpadding=1>")
      response.write("<tr>")
      if instr(cols,"F") > 0 then
        response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Name&nbsp;</b></font></td>")
      else
        response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Last&nbsp;</b></font></td>")
        response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;First&nbsp;</b></font></td>")
        response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Middle&nbsp;</b></font></td>")
      end if
      if instr(cols,"T") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Program&nbsp;(Track)&nbsp;</b></font></td>")
      if instr(cols,"H") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Harvard&nbsp;ID&nbsp;</b></font></td>")
      if instr(cols,"N") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;SSN&nbsp;</b></font></td>")
      if instr(cols,"B") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;DOB&nbsp;</b></font></td>")
      if instr(cols,"M") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Matric&nbsp;Dt&nbsp;</b></font></td>")
      if instr(cols,"G") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Exp&nbsp;Grad&nbsp;Dt&nbsp;</b></font></td>")
      if instr(cols,"Y") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;YOS&nbsp;</b></font></td>")
      if instr(cols,"X") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Gender&nbsp;</b></font></td>")
      if instr(cols,"E") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Ethnicity&nbsp;</b></font></td>")
      if instr(cols,"C") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Citizenship&nbsp;</b></font></td>")
      if instr(cols,"S") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Society&nbsp;</b></font></td>")
      if instr(cols,"L") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Email&nbsp;</b></font></td>")
	  if instr(cols,"I") > 0 then
	  	response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Hispanic&nbsp;</b></font></td>")
		response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;American Indian or Alaskan Native&nbsp;</b></font></td>")
		response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Asian&nbsp;</b></font></td>")
		response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Black or African American&nbsp;</b></font></td>")
		response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;Native Hawaiian or Pacific Islander&nbsp;</b></font></td>")
		response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;White&nbsp;</b></font></td>")
	  end if
	  if instr(cols,"D") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;DentPin&nbsp;</b></font></td>")
	  if instr(cols,"A") > 0 then response.write("<td align=center bgcolor=EEEEEE><font face='arial,helvetica' size=-1><b>&nbsp;AAMC&nbsp;ID&nbsp;</b></font></td>")
      response.write("</tr>")
      while not rs.eof
        huid = rs("harvard_id")
        if isnull(huid) then huid = ""
        response.write("<tr bgcolor=FFFFFF onMouseOver=""this.bgColor='#FFFF99'"" onMouseOut=""this.bgColor='#FFFFFF'"" onClick=""parent.document.location='p_summary.asp?id=" & rs("student_instance_id") & "'"">")
        if instr(cols,"F") > 0 then
          response.write("<td align=left nowrap=true><font face='arial,helvetica' size=-1>&nbsp;" & rs("full_name") & "&nbsp;</font></td>")
        else
          response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("last_name") & "&nbsp;</font></td>")
          response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("first_name") & "&nbsp;</font></td>")
          response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("middle_name") & "&nbsp;</font></td>")
        end if
        if instr(cols,"T") > 0 then response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("program_name") & " (" & rs("track_name") & ")&nbsp;</font></td>")
        if instr(cols,"H") > 0 then response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & huid & "&nbsp;</font></td>")
        if instr(cols,"N") > 0 then response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("ssn") & "&nbsp;</font></td>")
        if instr(cols,"B") > 0 then response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("date_of_birth") & "&nbsp;</font></td>")
        if instr(cols,"M") > 0 then response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & FormatSmallDate(rs("matric_date")) & "&nbsp;</font></td>")
        if instr(cols,"G") > 0 then response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & FormatSmallDate(rs("expected_grad_date")) & "&nbsp;</font></td>")
        if instr(cols,"Y") > 0 then response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("year_of_study_name") & "&nbsp;</font></td>")
        if instr(cols,"X") > 0 then response.write("<td align=center><font face='arial,helvetica' size=-1>&nbsp;" & rs("gender") & "&nbsp;</font></td>")
        if instr(cols,"E") > 0 then response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("ethnicity") & "&nbsp;</font></td>")
        if instr(cols,"C") > 0 then response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("citizenship") & "&nbsp;</font></td>")
        if instr(cols,"S") > 0 then response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("society") & "&nbsp;</font></td>")
        if instr(cols,"L") > 0 then response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("email_address") & "&nbsp;</font></td>")
		if instr(cols,"I") > 0 then
			response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("hispanic_flag") & "&nbsp;</font></td>")
			response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("amind_flag") & "&nbsp;</font></td>")
			response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("asian_flag") & "&nbsp;</font></td>")
			response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("black_flag") & "&nbsp;</font></td>")
			response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("pacif_flag") & "&nbsp;</font></td>")
			response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("white_flag") & "&nbsp;</font></td>")
		end if
		if instr(cols,"D") > 0 then response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("dentpin") & "&nbsp;</font></td>")
		if instr(cols,"A") > 0 then response.write("<td align=left><font face='arial,helvetica' size=-1>&nbsp;" & rs("amcas_id") & "&nbsp;</font></td>")
        response.write("</tr>")
        rs.movenext
      wend
      rs.close
      response.write("</table>")
      response.write("</td></tr></table>")
    end if
    response.write("</center>")
    response.write("</font>")
    response.write("</td></tr></table>")
    response.write("</body></html>")
  end if
end sub




''***************************************************
''**                    BILLING                    **
''***************************************************




sub BillingShell
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Term Billing"
  PageHeader
%>
<script>
function DoSave() {
  document.details.DoSave();
}
</script>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left valign=middle height=100%>
<iframe name="details" src="p_sets.asp?task=viewbilling&id=<%=id%>" scrolling=auto width=100% height=100% frameborder=0></iframe>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="DoSave();" name="bsave" value=" Save ">
<img src="images/spacer.gif" border=0 width=10 height=1>
<input style="width:65px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp?task=menu&id=<%=id%>');" name="bcancel" value=" Back ">
</nobr></td>
<%
PageFooter
end sub


sub ViewBilling
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
function DoSave() {
  if (confirm('Are you sure you want to run this batch process?')) {
    document.dataform.target = "_parent";
    document.dataform.task.value = "savebilling";
    document.dataform.submit();
  }
}
</script>
</head>
<body bgcolor="F9F9F9" text="000000" link="000000" vlink="000000" alink="000000">
<%
  sql = "select set_name, set_size from student_sets where student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_name = rs("set_name")
    set_size = rs("set_size")
  end if
  rs.close
  if isnull(set_name) then set_name = ""
  if set_name = "" then  set_name = "Untitled Set"
  if isnull(set_size) then set_size = ""
  if set_size = "" then set_size = "0"
  if set_size = "1" then
    set_size = set_size & "&nbsp;student"
  else
    set_size = set_size & "&nbsp;students"
  end if
  response.write("<font size=-1>")
  response.write("<center>")
  response.write("<table border=0 cellspacing=0 cellpadding=0 height=100% ><tr>")
  response.write("<td><form name=dataform method=""POST"" action=""p_sets.asp""><input type=hidden name=""task"" value=""viewbilling""><input type=hidden name=""id"" value=""" & id & """></td>")
  response.write("<td valign=middle>")

  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign=middle>")
  response.write("<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>")
  response.write("<font face='arial,helvetica'>")
  response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% ><tr>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' color=666666><b>" & set_name & "</b></font></td>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;" & set_size & "</font></td>")
  response.write("</tr></table>")
  response.write("<img src=""images/spacer.gif"" border=0 width=300 height=10><br>")
  response.write("<center>")

  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=AAAAAA>")
  response.write("<table border=0 cellspacing=1 cellpadding=2><tr><td align=center valign=middle bgcolor=EEEEEE>")
  ischecked = ""
  if request("tuition_count") = "Y" then ischecked = "checked"
  response.write("<font face='arial,helvetica' size=-1><input type=checkbox name=""tuition_count"" value=""Y"" "&ischecked&">add one to tuition count&nbsp;</font>")
  response.write("</td></tr></table>")
  response.write("</td></tr></table>")

  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")

  response.write("<table border=0 cellspacing=0 cellpadding=1><tr>")
  response.write("<td bgcolor=AAAAAA>")
  response.write("<table border=0 cellspacing=0 cellpadding=5>")
  response.write("<tr>")
  response.write("<td align=right bgcolor=DDDDDD><font face='arial,helvetica' size=-1><b>Term&nbsp;Filter</b>:</font></td>")
  response.write("<td align=left bgcolor=DDDDDD><select name=""term_id"" style=""width:150px"" onchange=""document.dataform.submit();"">")
  term_id = request("term_id")
  if term_id = "" then
	  term_id = "74"
  end if
  sql = "select * from terms as t order by t.start_date desc"
  rs.open sql,conn,1,1
  while not rs.eof
    if (term_id = "") and (date >= cdate(rs("start_date"))) and (date < cdate(rs("end_date"))) then
      term_id = rs("term_id")
    end if
    if cstr(term_id) = cstr(rs("term_id")) then
      isselected = " selected"
    else
      isselected = ""
    end if
    response.write("<option value=""" & rs("term_id") & """" & isselected & ">" & rs("term_name") & "</option>")
    rs.movenext
  wend
  rs.close
  response.write("</select>")
  response.write("</td>")
  response.write("</tr>")
  response.write("</table>")
  response.write("</td>")
  response.write("</tr></table>")

  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")

  sql = "select possible_charge_id, zone_name, charge_item_name, charge_type, amount from term_billing_charge_items where term_id = " & checkstring(term_id,50)
  if session("zone") <> "" then sql = sql & " and zone_name = " & checkstring(session("zone"),50)
  sql = sql & " order by zone_name, charge_item_name, amount, charge_type"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write("<font size=-1><i>There are no possible charge items for this term.</i><br></font>")
    response.write("<img src=""images/spacer.gif"" border=0 width=1 height=5><br>")
  else
    response.write("<font size=-1>Select the items you would like to charge:</font><br></font>")
    response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")
    response.write("<table border=0 cellspacing=0 cellpadding=1><tr>")
    response.write("<td bgcolor=AAAAAA>")
    response.write("<table border=0 cellspacing=0 cellpadding=5>")
    response.write("<tr>")
    response.write("<td align=left bgcolor=DDDDDD><font face='arial,helvetica' size=-1>")
    n = 0
    while not rs.eof
      n = n + 1
      charge_name = rs("charge_item_name")
      if isnull(charge_name) then charge_name = ""
      if charge_name = "" then charge_name = "Untitled Charge"
      if rs("charge_type") <> "" then charge_name = charge_name & " (" & rs("charge_type") & ")"
      if not isnull(rs("amount")) then charge_name = charge_name & " $" & rs("amount")
      if session("zone") = "" then
        charge_name = "[" & rs("zone_name") & "] " & charge_name
      end if
      response.write("<input type=checkbox name=""possible_charge_id_" & n & """ value=""" & rs("possible_charge_id") & """>" & charge_name & "<br>")
      rs.movenext
    wend
    rs.close
    response.write("<input type=hidden name=""rows"" value=""" & n & """>")
    response.write("</font></td>")
    response.write("</tr>")
    response.write("</table>")
    response.write("</td>")
    response.write("</tr></table>")
  end if

  response.write("</font>")
  response.write("</center>")
  response.write("</td></tr></table>")
  response.write("</td></tr></table>")

  response.write("</td>")
  response.write("<td></form></td>")
  response.write("</tr></table>")
  response.write("</body></html>")
end sub


sub SaveBilling
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Term Billing"
  n = request("rows")
  charge_items = ""
  if n <> "" then
    n = cint(n)
    for i = 1 to n
      if request("possible_charge_id_" & i) <> "" then
        if charge_items <> "" then charge_items = charge_items & ","
        charge_items = charge_items & request("possible_charge_id_" & i)
      end if
    next
  end if
  set_query = "select 0 student_instance_id where 1=0"
  sql = "select set_query from student_sets where (set_query is not null) and student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_query = rs("set_query")
  end if
  rs.close
  if charge_items <> "" then
    sql = ""
    sql = sql & " DECLARE @ErrorSave INT"
    sql = sql & " SET @ErrorSave = 0"
    sql = sql & " BEGIN TRANSACTION"
    sql = sql & " insert into term_billing (student_instance_id, possible_charge_id, created_by, created_dt, updated_by, updated_dt) "
    sql = sql & " select x.student_instance_id, x.possible_charge_id, "
    sql = sql & checkstring(session("user_id"),50) & "," & checkstring(now,50) & ","
	sql = sql & checkstring(session("user_id"),50) & "," & checkstring(now,50)
	sql = sql & " from "
	sql = sql & " ( "
	sql = sql & " select c.possible_charge_id, t.student_instance_id "
	sql = sql & " from term_billing_charge_items c, (" & set_query & ") t "
	sql = sql & " where c.possible_charge_id in (" & charge_items & ") "
	sql = sql & " ) x left outer join term_billing b on x.possible_charge_id = b.possible_charge_id and x.student_instance_id = b.student_instance_id "
	sql = sql & " where b.term_billing_id is null "
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"
    if request("tuition_count") = "Y" then
      sql = sql & " update student_instance set tuition_count = (coalesce(tuition_count,0)+1) "
      sql = sql & " , updated_by = " & checkstring(session("user_id"),50) & ", updated_dt = " & checkstring(now,50)
      sql = sql & " where student_instance_id in (" & set_query & ") "
      sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"
    end if
    sql = sql & " IF (@ErrorSave = 0)"
    sql = sql & "   COMMIT TRANSACTION"
    sql = sql & " ELSE"
    sql = sql & "   ROLLBACK TRANSACTION"
    sql = sql & " SELECT @ErrorSave AS theError"
    theError = 1
    rs.open sql,conn,1,1
    do until (rs is nothing)
      if (rs.state = 1) then
        if (not rs.eof) then
          theError = rs("theError")
        end if
      end if
      set rs = rs.NextRecordset()
    loop
    set rs = Server.CreateObject("ADODB.RecordSet")
    if cint(theError) = 0 then
      SaveComplete("The selected items have been charged to all students in this set.")
    else
      SaveComplete("<b>An error has occurred. No changes were made.</b>")
    end if
  else
    if request("tuition_count") = "Y" then
      sql = "update student_instance set tuition_count = (coalesce(tuition_count,0)+1) "
      sql = sql & " , updated_by = " & checkstring(session("user_id"),50) & ", updated_dt = " & checkstring(now,50)
      sql = sql & " where student_instance_id in (" & set_query & ") "
      conn.execute(sql)
      SaveComplete("The tuition counts have been updated.")
    else
      SaveComplete("No changes were made to the database.")
    end if
  end if
end sub




''***************************************************
''**                    ADVISORS                   **
''***************************************************




sub ViewAdvisors
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Advisor Assignment"
  PageHeader
%>
<script>
function DoSave() {
  if (document.dataform.role_id.selectedIndex == 0) {
    alert('Please select an advisor role.');
  } else if (document.dataform.advisor_username.value == '') {
    alert('Please select an advisor.');
  } else if (confirm('Are you sure you want to run this batch process?')) {
    document.dataform.submit();
  }
}
function PickPerson() {
  document.dataform.advisor_username.value='';
  document.dataform.advisor_name.value='';
  document.dataform.advisor_phone.value='';
  document.dataform.advisor_email.value='';
  window.open("g_person.asp?fi=advisor_username&fn=advisor_name&fp=advisor_phone&fe=advisor_email","PersonWindow","width=550,height=400,menubar,scrollbars,resizable,copyhistory=no");
}
</script>
<!--#include file ="js_cal.asp"-->
<!--#include file ="formval.asp"-->
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=center valign=middle height=100%>
<%
  sql = "select set_name, set_size from student_sets where student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_name = rs("set_name")
    set_size = rs("set_size")
  end if
  rs.close
  if isnull(set_name) then set_name = ""
  if set_name = "" then  set_name = "Untitled Set"
  if isnull(set_size) then set_size = ""
  if set_size = "" then set_size = "0"
  if set_size = "1" then
    set_size = set_size & "&nbsp;student"
  else
    set_size = set_size & "&nbsp;students"
  end if
  response.write("<table border=0 cellspacing=0 cellpadding=0 height=100% ><tr>")
  response.write("<td><form name=dataform method=""POST"" action=""p_sets.asp"" target=""_parent""><input type=hidden name=""task"" value=""saveadvisors""><input type=hidden name=""id"" value=""" & id & """></td>")
  response.write("<td valign=middle>")

  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign=middle>")
  response.write("<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>")
  response.write("<font face='arial,helvetica'>")
  response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% ><tr>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' color=666666><b>" & set_name & "</b></font></td>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;" & set_size & "</font></td>")
  response.write("</tr></table>")
  response.write("<img src=""images/spacer.gif"" border=0 width=300 height=15><br>")
  response.write("<center>")
%>

<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Role:&nbsp;</font></td>
<td align=left><select name="role_id" style="width:200px"><option value=""></option>
<%
sql = "select * from advisor_roles order by role_name"
rs.open sql,conn,1,1
while not rs.eof
  response.write("<option value=""" & rs("role_id") & """" & isselected & ">" & rs("role_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Start&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="start_date" value="" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);"><%DrawCal "dataform.start_date", ""%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;End&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="end_date" value="" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);"><%DrawCal "dataform.end_date", ""%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Advisor:&nbsp;</font></td>
<td align=left colspan=5>
<input type=hidden name="advisor_username" value="">
<input type=text name="advisor_name" value="" size=20 style="width:175px" contenteditable=false>
<input type=text name="advisor_phone" value="" size=20 style="width:100px" contenteditable=false>
<input type=text name="advisor_email" value="" size=20 style="width:215px" contenteditable=false>
<a href="JavaScript:PickPerson();"><img src="images/search.gif" border=0></a>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=5><input type=text name="notes" size=15 style="width:525px"</td>
</tr>
</table>

<%
  response.write("</center>")
  response.write("</font>")
  response.write("</td></tr></table>")
  response.write("</td></tr></table>")

  response.write("</td>")
  response.write("<td></form></td>")
  response.write("</tr></table>")
%>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="DoSave();" name="bsave" value=" Save ">
<img src="images/spacer.gif" border=0 width=10 height=1>
<input style="width:65px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp?task=menu&id=<%=id%>');" name="bcancel" value=" Back ">
</nobr></td>
<%
PageFooter
end sub



sub SaveAdvisors
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Term Billing"
  set_query = "select 0 student_instance_id where 1=0"
  sql = "select set_query from student_sets where (set_query is not null) and student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_query = rs("set_query")
  end if
  rs.close
  sql = "insert into advisors (advisor_username, student_instance_id, person_id, role_id, start_date, end_date, notes, created_by, created_dt, updated_by, updated_dt) "
  sql = sql & " select x.advisor_username, x.student_instance_id, x.person_id, x.role_id, "
  sql = sql & checkstring2(request("start_date"),true) & "," & checkstring2(request("end_date"),true) & "," & checkstring(request("notes"),len(request("notes"))) & ","
  sql = sql & checkstring(session("user_id"),50) & "," & checkstring(now,50) & ","
  sql = sql & checkstring(session("user_id"),50) & "," & checkstring(now,50)
  sql = sql & " from "
  sql = sql & " ( "
  sql = sql & " select " & checkstring(request("advisor_username"),50) & " as advisor_username, i.student_instance_id, i.person_id, " & checkstring(request("role_id"),50) & " as role_id "
  sql = sql & " from student_instance i, (" & set_query & ") t "
  sql = sql & " where i.student_instance_id = t.student_instance_id "
  sql = sql & " ) x left outer join advisors a on x.role_id = a.role_id and (x.student_instance_id = a.student_instance_id or x.person_id = a.person_id) "
  sql = sql & " where a.advisor_id is null "
  conn.execute(sql)
  SaveComplete("The selected advisor has been assigned to all students in this set.")
end sub







''***************************************************
''**                  REG HISTORY                  **
''***************************************************




sub ViewRegStatus
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Registration Status"
  PageHeader
%>
<script>
function DoSave() {
  if (confirm('Are you sure you want to run this batch process?')) {
    document.dataform.submit();
  }
}
function SetGradStatus1() {
  if (document.dataform.reg_status_id.options[document.dataform.reg_status_id.selectedIndex].value == '42') {
    document.dataform.set_graduated.checked = true;
  } else {
    document.dataform.set_graduated.checked = false;
  }
}
function SetGradStatus2() {
  if (document.dataform.set_graduated.checked) {
    for (i = 0; i < document.dataform.reg_status_id.options.length; i++) {
      if (document.dataform.reg_status_id.options[i].value == '42') {
        document.dataform.reg_status_id.selectedIndex = i;
      }
    }
  }
}
function SetRegDate(which) {
  if (which == 1) {
    document.dataform.regstart[0].checked = true;
  }
  if (which == 2) {
    document.dataform.regstart[1].checked = true;
  }
  if (which == 3) {
    document.dataform.regend[0].checked = true;
  }
  if (which == 4) {
    document.dataform.regend[1].checked = true;
  }
}
</script>
<!--#include file ="js_cal.asp"-->
<!--#include file ="formval.asp"-->
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=center valign=middle height=100%>
<%
  sql = "select set_name, set_size from student_sets where student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_name = rs("set_name")
    set_size = rs("set_size")
  end if
  rs.close
  if isnull(set_name) then set_name = ""
  if set_name = "" then  set_name = "Untitled Set"
  if isnull(set_size) then set_size = ""
  if set_size = "" then set_size = "0"
  if set_size = "1" then
    set_size = set_size & "&nbsp;student"
  else
    set_size = set_size & "&nbsp;students"
  end if
  response.write("<table border=0 cellspacing=0 cellpadding=0 height=100% ><tr>")
  response.write("<td><form name=dataform method=""POST"" action=""p_sets.asp"" target=""_parent""><input type=hidden name=""task"" value=""saveregstatus""><input type=hidden name=""id"" value=""" & id & """></td>")
  response.write("<td valign=middle>")

  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign=middle>")
  response.write("<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>")
  response.write("<font face='arial,helvetica'>")
  response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% ><tr>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' color=666666><b>" & set_name & "</b></font></td>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;" & set_size & "</font></td>")
  response.write("</tr></table>")
  response.write("<img src=""images/spacer.gif"" border=0 width=300 height=15><br>")
%>
<font size=-1>

<b>Student Instance</b><br>
<img src="images/spacer.gif" border=0 width=300 height=5><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Status:&nbsp;</font></td>
<td align=left valign=middle><select name="reg_status_id" style="width:250px" onChange="SetGradStatus1();document.dataform.status_date.value='<%=date%>';"><option value="">(no change)</option>
<%
sql = "select * from reg_status where zone_name = '' or zone_name = '" & session("zone") & "' order by reg_status_name"
rs.open sql,conn,1,1
while not rs.eof
  response.write("<option value=""" & rs("reg_status_id") & """>" & rs("reg_status_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Status&nbsp;Dt:&nbsp;</font></td>
<td align=left valign=middle><input type=text name="status_date" value="" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);"><%DrawCal "dataform.status_date", ""%></td>
</tr>
</table>
<img src="images/spacer.gif" border=0 width=1 height=2><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Year of Study:&nbsp;</font></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1><input type=radio name="yos" value="" checked>no change</font></nobr></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input type=radio name="yos" value="adv">advance ID by one</font></nobr></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input type=radio name="yos" value="val">set to <select name="year_of_study_id" style="width:60px" onChange="document.dataform.yos[2].checked=true"><option value="">(null)</option>
<%
sql = "select * from year_of_study order by year_of_study_id"
rs.open sql,conn,1,1
while not rs.eof
  response.write("<option value=""" & rs("year_of_study_id") & """>" & rs("year_of_study_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</font></nobr></td>
</tr>
<tr><td colspan=4><img src="images/spacer.gif" border=0 width=1 height=2></td></tr>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Years in Prgm:&nbsp;</font></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1><input type=radio name="yip" value="" checked>no change</font></nobr></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input type=radio name="yip" value="adv">advance by one</font></nobr></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;<input type=radio name="yip" value="val">set to <input type=text name="years_in_program" value="" size=10 maxlength=5 style="width:80px" onChange="document.dataform.yip[2].checked=true"></font></nobr></td>
</tr>
</table>

<img src="images/spacer.gif" border=0 width=1 height=10><br>

<b>Reg Hist</b><br>
<img src="images/spacer.gif" border=0 width=1 height=10><br>
<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Action:&nbsp;</font></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1>
<input type=radio name="reg" value="">no change
&nbsp;&nbsp;<input type=radio name="reg" value="new">create new
&nbsp;&nbsp;<input type=radio name="reg" value="prog" checked>create new if program keeps reg hist
</font></nobr></td>
</tr>
<tr><td colspan=2><img src="images/spacer.gif" border=0 width=1 height=5></td></tr>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Start&nbsp;Dt:&nbsp;</font></td>
<td align=left valign=middle><table border=0 cellspacing=0 cellpadding=0><tr>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1><input type=radio name="regstart" value="manual" checked>enter date:&nbsp;</font></nobr></td>
<td align=left valign=middle><input type=text name="reg_start_date" value="<%=date%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);SetRegDate(1);"><%DrawCal "dataform.reg_start_date", "SetRegDate(1);"%></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;&nbsp;&nbsp;<input type=radio name="regstart" value="auto">use start of term&nbsp;</font></nobr></td>
</tr></table></td>
</tr>
<tr><td colspan=2><img src="images/spacer.gif" border=0 width=1 height=1></td></tr>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;&nbsp;End&nbsp;Dt:&nbsp;</font></td>
<td align=left valign=middle><table border=0 cellspacing=0 cellpadding=0><tr>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1><input type=radio name="regend" value="manual">enter date:&nbsp;</font></nobr></td>
<td align=left valign=middle><input type=text name="reg_end_date" value="<%=date%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);SetRegDate(3);"><%DrawCal "dataform.reg_end_date", "SetRegDate(3);"%></td>
<td align=left valign=middle><nobr><font face='arial,helvetica' size=-1>&nbsp;&nbsp;&nbsp;&nbsp;<input type=radio name="regend" value="auto" checked>use end of term&nbsp;</font></nobr></td>
</tr></table></td>
</tr>
<tr><td colspan=2><img src="images/spacer.gif" border=0 width=1 height=5></td></tr>
<tr>
<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Term:&nbsp;</font></td>
<td align=left><select name="term_id" style="width:105px"><option value=""></option>
<%
sql = "select * from terms order by start_date desc"
rs.open sql,conn,1,1
while not rs.eof
  is_selected = ""
  if date >= rs("start_date") and date <= rs("end_date") then is_selected = " selected"
  response.write("<option value=""" & rs("term_id") & """" & is_selected & ">" & rs("term_name") & "</option>")
  rs.movenext
wend
rs.close
%>
</select>
</td>
</tr>

</table>

<img src="images/spacer.gif" border=0 width=1 height=10><br>

<b>Prematriculants</b><br>
<img src="images/spacer.gif" border=0 width=1 height=1><br>
<nobr>&nbsp;&nbsp;<input type=checkbox name="set_matric_date" value="Y">copy <u>reg hist start date</u> to matric_date, delete prematric reg hist, add degree</nobr><br>

<img src="images/spacer.gif" border=0 width=1 height=10><br>

<b>Graduation</b><br>
<img src="images/spacer.gif" border=0 width=1 height=1><br>
<nobr>&nbsp;&nbsp;<input type=checkbox name="set_graduated" value="Y" onClick="SetGradStatus2()">copy <u>reg hist start date</u> to actual_grad_date, set grad flag, update degree</nobr><br>


</font>
<%
  response.write("</font>")
  response.write("</td></tr></table>")
  response.write("</td></tr></table>")

  response.write("</td>")
  response.write("<td></form></td>")
  response.write("</tr></table>")
%>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="DoSave();" name="bsave" value=" Save ">
<img src="images/spacer.gif" border=0 width=10 height=1>
<input style="width:65px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp?task=menu&id=<%=id%>');" name="bcancel" value=" Back ">
</nobr></td>
<%
PageFooter
end sub



sub SaveRegStatus
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Registration Status"

  set_query = "select 0 student_instance_id where 1=0"
  sql = "select set_query from student_sets where (set_query is not null) and student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_query = rs("set_query")
  end if
  rs.close

  sql = ""
  sql = sql & " DECLARE @ErrorSave INT"
  sql = sql & " SET @ErrorSave = 0"
  sql = sql & " BEGIN TRANSACTION"

  sql = sql & " DECLARE @student_instance_list TABLE (student_instance int) "
  sql = sql & " INSERT INTO @student_instance_list " & set_query
  sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"

  set_query = " SELECT student_instance FROM @student_instance_list "

  q_yos = "i.year_of_study_id"
  if request("yos") = "adv" then
    q_yos = "(coalesce(i.year_of_study_id,0)+1)"
  end if
  if request("yos") = "val" then
    q_yos = checkstring2(request("year_of_study_id"),true)
  end if

  q_yip = "i.years_in_program"
  if request("yip") = "adv" then
    q_yip = "(coalesce(i.years_in_program,0) + 1)"
  elseif request("yip") = "val" then
    q_yip = checkstring2(request("years_in_program"),true)
  end if

  q_regstart = "null"
  if request("regstart") = "manual" then
    q_regstart = checkstring2(request("reg_start_date"),true)
  end if
  if request("regstart") = "auto" then
    q_regstart = "dbo.GetRegHistDate('S'," & checkstring2(request("term_id"),true) & ",null,i.program_id," & q_yos & ") "
  end if

  q_regend = null
  if request("regend") = "manual" then
    q_regend = checkstring2(request("reg_end_date"),true)
  end if
  if request("regend") = "auto" then
    q_regend = "dbo.GetRegHistDate('E'," & checkstring2(request("term_id"),true) & ",null,i.program_id," & q_yos & ") "
  end if

  q_status = "i.reg_status_id"
  if request("reg_status_id") <> "" then
    q_status = checkstring2(request("reg_status_id"),true)
  end if

  if request("reg") = "new" or request("reg") = "prog" then

    if request("set_matric_date") = "Y" then
      sql = sql & " DELETE FROM student_reg_hist WHERE reg_status_id = '15' AND student_instance_id IN (" & set_query & ") "
    end if

    sql = sql & " UPDATE r SET r.effective_end_date = " & q_regstart
    sql = sql & " , r.updated_by = " & checkstring(session("user_id"),50) & ", r.updated_dt = " & checkstring(now,50)
    sql = sql & " FROM student_reg_hist r, student_instance i, programs g "
    sql = sql & " WHERE r.student_instance_id = i.student_instance_id AND i.program_id = g.program_id "
    if request("reg") = "prog" then sql = sql & " AND g.track_reg_hist = 'Y' "
    sql = sql & " AND r.effective_end_date > " & q_regstart
    sql = sql & " AND i.student_instance_id IN (" & set_query & ") "
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"

    sql = sql & " INSERT INTO student_reg_hist (student_instance_id,term_id,reg_status_id,person_id,effective_start_date,effective_end_date,year_of_study_id,reg_status_name,notes,created_by,created_dt,updated_by,updated_dt) "
    sql = sql & " SELECT i.student_instance_id, " & checkstring2(request("term_id"),true) & ", " & q_status & ", i.person_id, " & q_regstart & ", " & q_regend & ", " & q_yos & ", s.reg_status_name, null, "
    sql = sql & checkstring(session("user_id"),50) & "," & checkstring(now,50) & "," & checkstring(session("user_id"),50) & "," & checkstring(now,50)
    sql = sql & " FROM student_instance i, programs g, reg_status s "
    sql = sql & " WHERE i.program_id = g.program_id AND s.reg_status_id = " & q_status
    if request("reg") = "prog" then sql = sql & " AND g.track_reg_hist = 'Y' "
    sql = sql & " AND i.student_instance_id IN (" & set_query & ") "
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"

  end if

  if request("set_matric_date") = "Y" then
    sql = sql & " UPDATE i SET i.matric_date = " & q_regstart
    sql = sql & " , i.updated_by = " & checkstring(session("user_id"),50) & ", i.updated_dt = " & checkstring(now,50)
    sql = sql & " FROM student_instance i "
    sql = sql & " WHERE i.student_instance_id IN (" & set_query & ") "
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"

    sql = sql & " INSERT INTO student_degrees (person_id,degree_id,student_instance_id,degree_status_id,program_track_id,institution_id,start_date,created_by,created_dt,updated_by,updated_dt) "
    sql = sql & " SELECT i.person_id, d.degree_id, i.student_instance_id, '1', i.program_track_id, null, " & q_regstart
    sql = sql & " ," & checkstring(session("user_id"),50) & ", " & checkstring(now,50) & ", " & checkstring(session("user_id"),50) & ", " & checkstring(now,50)
    sql = sql & " FROM student_instance i, program_degrees d WHERE d.program_id = i.program_id AND i.student_instance_id IN (" & set_query & ") "
    sql = sql & " AND NOT EXISTS (SELECT * FROM student_degrees sd WHERE sd.student_instance_id = i.student_instance_id) "
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"
  end if

  if request("set_graduated") = "Y" then
    sql = sql & " UPDATE i SET i.graduated = 'Y', i.actual_grad_date = " & q_regstart
    sql = sql & " , i.updated_by = " & checkstring(session("user_id"),50) & ", i.updated_dt = " & checkstring(now,50)
    sql = sql & " FROM student_instance i "
    sql = sql & " WHERE i.student_instance_id IN (" & set_query & ") "
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"

    sql = sql & " UPDATE d SET d.degree_status_id = '2', d.end_date = " & q_regstart & ", d.completed_date = " & q_regstart
    sql = sql & " , d.updated_by = " & checkstring(session("user_id"),50) & ", d.updated_dt = " & checkstring(now,50)
    sql = sql & " FROM student_instance i, student_degrees d "
    sql = sql & " WHERE i.student_instance_id = d.student_instance_id "
    sql = sql & " AND i.student_instance_id IN (" & set_query & ") "
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"
  end if

  if request("reg_status_id") <> "" or request("status_date") <> "" or request("yos") <> "" or request("yip") <> "" then
    sql = sql & " UPDATE i SET i.year_of_study_id = " & q_yos & ", i.years_in_program = " & q_yip
    sql = sql & " , i.reg_status_id = " & q_status
    if request("status_date") <> "" then
      sql = sql & ", i.reg_status_date = " & checkstring2(request("status_date"),true)
    end if
    sql = sql & " , i.updated_by = " & checkstring(session("user_id"),50) & ", i.updated_dt = " & checkstring(now,50)
    sql = sql & " FROM student_instance i WHERE i.student_instance_id IN (" & set_query & ") "
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"
  end if

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
  set rs = Server.CreateObject("ADODB.RecordSet")
  if cint(theError) = 0 then
    SaveComplete("The selected items have been saved.")
  else
    SaveComplete("<b>An error has occurred. No changes were made.</b>")
  end if
end sub






''***************************************************
''**                  REQUIREMENTS                 **
''***************************************************



sub RequirementsShell
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Requirements"
  PageHeader
%>
<script>
function DoSave() {
  document.details.DoSave();
}
</script>
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=left valign=middle height=100%>
<iframe name="details" src="p_sets.asp?task=viewrequirements&id=<%=id%>" scrolling=auto width=100% height=100% frameborder=0></iframe>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="DoSave();" name="bsave" value=" Save ">
<img src="images/spacer.gif" border=0 width=10 height=1>
<input style="width:65px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp?task=menu&id=<%=id%>');" name="bcancel" value=" Back ">
</nobr></td>
<%
PageFooter
end sub




sub ViewRequirements
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
function DoSave() {
  if (confirm('Are you sure you want to run this batch process?')) {
    document.dataform.submit();
  }
}
</script>
</head>
<body bgcolor="F9F9F9" text="000000" link="000000" vlink="000000" alink="000000">
<!--#include file ="js_cal.asp"-->
<!--#include file ="formval.asp"-->
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=center valign=middle height=100%>
<%
  sql = "select set_name, set_size from student_sets where student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_name = rs("set_name")
    set_size = rs("set_size")
  end if
  rs.close
  if isnull(set_name) then set_name = ""
  if set_name = "" then  set_name = "Untitled Set"
  if isnull(set_size) then set_size = ""
  if set_size = "" then set_size = "0"
  if set_size = "1" then
    set_size = set_size & "&nbsp;student"
  else
    set_size = set_size & "&nbsp;students"
  end if
  response.write("<table border=0 cellspacing=0 cellpadding=0 height=100% ><tr>")
  response.write("<td><form name=dataform method=""GET"" action=""p_sets.asp"" target=""_parent""><input type=hidden name=""task"" value=""saverequirements""><input type=hidden name=""id"" value=""" & id & """></td>")
  response.write("<td valign=middle>")

  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign=middle>")
  response.write("<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>")
  response.write("<font face='arial,helvetica'>")
  response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% ><tr>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' color=666666><b>" & set_name & "</b></font></td>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;" & set_size & "</font></td>")
  response.write("</tr></table>")
  response.write("<img src=""images/spacer.gif"" border=0 width=300 height=15><br>")
  response.write("<center>")

  response.write("<font size=-1>Select a save type and enter optional values:</font><br>")

  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")

  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=AAAAAA>")
  response.write("<table border=0 cellspacing=1 cellpadding=2>")

  response.write("<tr><td align=center valign=middle bgcolor=EEEEEE>")
  response.write("<table border=0 cellspacing=0 cellpadding=0>")
  response.write("<tr><td align=left valign=top><nobr><font face='arial,helvetica' size=-1><input type=radio name=""savetype"" value=""nosave"" checked>Do not save if requirement exists</font></nobr></td></tr>")
  response.write("<tr><td align=left valign=top><nobr><font face='arial,helvetica' size=-1><input type=radio name=""savetype"" value=""overwrite"">Overwrite if requirement exists</font></nobr></td></tr>")
  response.write("<tr><td align=left valign=top><nobr><font face='arial,helvetica' size=-1><input type=radio name=""savetype"" value=""duplicate"">Make duplicate if requirement exists</font></nobr></td></tr>")
  response.write("</table>")
  response.write("<img src=""images/spacer.gif"" border=0 width=250 height=1><br>")
  response.write("</td></tr>")

  response.write("<tr><td align=center valign=middle bgcolor=EEEEEE>")
  response.write("<table border=0 cellspacing=0 cellpadding=0>")
  response.write("<tr>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>Completed:&nbsp;</font></td>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' size=-1><input type=checkbox name=""completed"" value=""Y"">check if completed</font></td>")
  response.write("</tr>")
  response.write("<tr>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>Completed&nbsp;Dt:&nbsp;</font></td>")
  response.write("<td align=left valign=middle><input type=text name=""completed_date"" value="""" size=15 maxlength=50 style=""width:80px"" onchange=""validateDate(this, false);"">")
  DrawCal "dataform.completed_date", ""
  response.write("</td>")
  response.write("</tr>")
  response.write("<tr>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>Expires&nbsp;Dt:&nbsp;</font></td>")
  response.write("<td align=left valign=middle><input type=text name=""expires_date"" value="""" size=15 maxlength=50 style=""width:80px"" onchange=""validateDate(this, false);"">")
  DrawCal "dataform.expires_date", ""
  response.write("</td>")
  response.write("</tr>")
  response.write("<tr>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>Status:&nbsp;</font></td>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' size=-1><input type=text name=""status"" value="""" size=15 maxlength=50 style=""width:150px""></font></td>")
  response.write("</tr>")
  response.write("<tr>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>Notes:&nbsp;</font></td>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' size=-1><input type=text name=""notes"" value="""" size=15 style=""width:150px""></font></td>")
  response.write("</tr>")
  response.write("</table>")
  response.write("</td></tr>")
  response.write("</table>")
  response.write("</td></tr></table>")

  response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")

  sql = "select requirement_id, requirement_name, zone_name from requirements "
  if session("zone") <> "" then
    sql = sql & " where (zone_name is null or zone_name = '' or zone_name = '" & session("zone") & "')"
  end if
  sql = sql & " order by requirement_name"
  rs.open sql,conn,1,1
  if rs.eof then
    response.write("<font size=-1><i>There are no available requirements.</i><br></font>")
    response.write("<img src=""images/spacer.gif"" border=0 width=1 height=5><br>")
  else
    response.write("<font size=-1>Select the requirements you would like to assign:</font><br></font>")
    response.write("<img src=""images/spacer.gif"" border=0 width=1 height=15><br>")
    response.write("<table border=0 cellspacing=0 cellpadding=1><tr>")
    response.write("<td bgcolor=AAAAAA>")
    response.write("<table border=0 cellspacing=0 cellpadding=5>")
    response.write("<tr>")
    response.write("<td align=left bgcolor=DDDDDD><font face='arial,helvetica' size=-1>")
    n = 0
    while not rs.eof
      n = n + 1
      response.write("<input type=checkbox name=""requirement_id_" & n & """ value=""" & rs("requirement_id") & """>" & rs("requirement_name") & "<br>")
      rs.movenext
    wend
    rs.close
    response.write("<input type=hidden name=""rows"" value=""" & n & """>")
    response.write("</font></td>")
    response.write("</tr>")
    response.write("</table>")
    response.write("</td>")
    response.write("</tr></table>")
  end if

  response.write("</center>")
  response.write("</font>")
  response.write("</td></tr></table>")
  response.write("</td></tr></table>")

  response.write("</td>")
  response.write("<td></form></td>")
  response.write("</tr></table>")
  response.write("</body></html>")
end sub



sub SaveRequirements
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Requirements"
  n = request("rows")
  req_list = ""
  if n <> "" then
    n = cint(n)
    for i = 1 to n
      if request("requirement_id_" & i) <> "" then
        if req_list <> "" then req_list = req_list & ","
        req_list = req_list & request("requirement_id_" & i)
      end if
    next
  end if
  set_query = "select 0 student_instance_id where 1=0"
  sql = "select set_query from student_sets where (set_query is not null) and student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_query = rs("set_query")
  end if
  rs.close
  if req_list <> "" then
    sql = ""
    sql = sql & " DECLARE @ErrorSave INT"
    sql = sql & " SET @ErrorSave = 0"
    sql = sql & " BEGIN TRANSACTION"

    sql = sql & " DECLARE @student_instance_list TABLE (student_instance int) "
    sql = sql & " INSERT INTO @student_instance_list " & set_query
    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"

    if request("savetype") = "overwrite" then
      sql = sql & " DELETE FROM student_requirements WHERE person_id In (SELECT person_id FROM student_instance WHERE student_instance_id in (" & set_query & ")) and requirement_id in (" & req_list & ") "
      sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"
    end if

    sql = sql & " insert into student_requirements (person_id, student_instance_id, requirement_id, completed, completed_date, expires_date, status, notes) "
    sql = sql & " select i.person_id, i.student_instance_id, r.requirement_id, "
    sql = sql & checkstring2(request("completed"),false) & "," & checkstring2(request("completed_date"),true) & ","
    if request("completed") = "Y" and request("completed_date") <> "" and request("expires_date") = "" then
      sql = sql & "(case when r.valid_days is not null then cast(convert(nvarchar(50),GetDate()+r.valid_days,1) as datetime) else null end), "
    else
      sql = sql & checkstring2(request("expires_date"),true) & ","
    end if
    sql = sql & checkstring2(request("status"),false) & "," & checkstring2(request("notes"),false)
    sql = sql & " from student_instance i, requirements r, (" & set_query & ") t "
    sql = sql & " where i.student_instance_id = t.student_instance_id and r.requirement_id in (" & req_list & ") "
    if request("savetype") = "nosave" then
      sql = sql & " and not exists (select x.student_requirement_id from student_requirements x where x.person_id = i.person_id and x.requirement_id = r.requirement_id) "
    end if

    sql = sql & " SET @ErrorSave = @ErrorSave + @@Error"

    sql = sql & " IF (@ErrorSave = 0)"
    sql = sql & "   COMMIT TRANSACTION"
    sql = sql & " ELSE"
    sql = sql & "   ROLLBACK TRANSACTION"
    sql = sql & " SELECT @ErrorSave AS theError"

    'response.write(sql)
    'response.end

    theError = 1
    rs.open sql,conn,1,1
    do until (rs is nothing)
      if (rs.state = 1) then
        if (not rs.eof) then
          theError = rs("theError")
        end if
      end if
      set rs = rs.NextRecordset()
    loop
    set rs = Server.CreateObject("ADODB.RecordSet")
    if cint(theError) = 0 then
      SaveComplete("The selected items have been assigned to all students in this set.")
    else
      SaveComplete("<b>An error has occurred. No changes were made.</b>")
    end if
  else
    SaveComplete("No changes were made to the database.")
  end if
end sub


sub RequirementsReport

end sub








''***************************************************
''**                 CORRESPONDENCE                **
''***************************************************




sub ViewCorrespondence
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Correspondence"
  PageHeader
%>
<script>
function DoSave() {
  if (document.dataform.letter_id.selectedIndex == 0) {
    alert('Please select letter type.');
  } else if (confirm('Are you sure you want to run this batch process?')) {
    document.dataform.submit();
  }
}
</script>
<!--#include file ="js_cal.asp"-->
<!--#include file ="formval.asp"-->
<table width=100% height=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=center valign=middle height=100%>
<%
  sql = "select set_name, set_size from student_sets where student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_name = rs("set_name")
    set_size = rs("set_size")
  end if
  rs.close
  if isnull(set_name) then set_name = ""
  if set_name = "" then  set_name = "Untitled Set"
  if isnull(set_size) then set_size = ""
  if set_size = "" then set_size = "0"
  if set_size = "1" then
    set_size = set_size & "&nbsp;student"
  else
    set_size = set_size & "&nbsp;students"
  end if
  response.write("<table border=0 cellspacing=0 cellpadding=0 height=100% ><tr>")
  response.write("<td><form name=dataform method=""POST"" action=""p_sets.asp"" target=""_parent""><input type=hidden name=""task"" value=""savecorrespondence""><input type=hidden name=""id"" value=""" & id & """></td>")
  response.write("<td valign=middle>")

  response.write("<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=CCCCCC valign=middle>")
  response.write("<table border=0 cellspacing=1 cellpadding=10><tr><td align=left bgcolor=EEEEEE>")
  response.write("<font face='arial,helvetica'>")
  response.write("<table border=0 cellspacing=0 cellpadding=0 width=100% ><tr>")
  response.write("<td align=left valign=middle><font face='arial,helvetica' color=666666><b>" & set_name & "</b></font></td>")
  response.write("<td align=right valign=middle><font face='arial,helvetica' size=-1>&nbsp;" & set_size & "</font></td>")
  response.write("</tr></table>")
  response.write("<img src=""images/spacer.gif"" border=0 width=300 height=15><br>")
  response.write("<center>")
%>

<table border=0 cellspacing=0 cellpadding=0>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Letter:&nbsp;</font></td>
<td align=left colspan=3><select name="letter_id" style="width:305px"><option value="">(delete)</option>
<%
sql = "SELECT * FROM letter_defs WHERE active='Y' order by letter_type, letter_name"
rs.open sql,conn,1,1
while not rs.eof
  response.write("<option value=""" & rs("letter_id") & """>" & rs("letter_name") & "</option>")
  rs.movenext
wend
rs.close

%>
</select>
</td>
<td align=right><font face='arial,helvetica' size=-1>Status:&nbsp;</font></td>
<td align=left><select name="status" style="width:100px">
<option value=""></option>
<option value="Sent">Sent</option>
<option value="Queued">Queued</option>
</select>
</td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Req&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="requested_date" value="<%=date%>" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);"><%DrawCal "dataform.requested_date_"&i, ""%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Sched&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="scheduled_date" value="" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);"><%DrawCal "dataform.scheduled_date", ""%></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Print&nbsp;Dt:&nbsp;</font></td>
<td align=left><input type=text name="printed_date" value="" size=15 maxlength=50 style="width:80px" onchange="validateDate(this, false);"><%DrawCal "dataform.printed_date", ""%></td>
</tr>
<tr>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Copies:&nbsp;</font></td>
<td align=left><input type=text name="copies" value="1" size=15 maxlength=50 style="width:80px"></td>
<td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Notes:&nbsp;</font></td>
<td align=left colspan=3><input type=text name="notes" value="" size=15 style="width:325px"></td>
</tr>
<tr><td align=right><font face='arial,helvetica' size=-1>&nbsp;&nbsp;Address&nbsp;Type:&nbsp;</font></td>
<td colspan="3">
<select name="addresstype" style="width:175px" id="Select1">
  <option value="Alternate Mailing">Alternate Mailing</option>
  <option value="Emergency contact">Emergency contact</option>
  <option value="Exclerk Certification">Exclerk Certification</option>
  <option value="Father Address">Father Address</option>
  <option value="Gradesheet address">Gradesheet address</option>
  <option value="Local/Mailing" selected>Local/Mailing</option>
  <option value="Mother Address">Mother Address</option>
  <option value="Permanent">Permanent</option>
</select>
</td></tr>
</table>

<%
  response.write("</center>")
  response.write("</font>")
  response.write("</td></tr></table>")
  response.write("</td></tr></table>")

  response.write("</td>")
  response.write("<td></form></td>")
  response.write("</tr></table>")
%>
</td>
</tr>
</table>
<%
PageMiddle
%>
<td align=center valign=middle><nobr>
<input style="width:65px; background-color:#CCCCCC" type=button onClick="DoSave();" name="bsave" value=" Save ">
<img src="images/spacer.gif" border=0 width=10 height=1>
<input style="width:65px; background-color:#CCCCCC" type=button onclick="GoPage('p_sets.asp?task=menu&id=<%=id%>');" name="bcancel" value=" Back ">
</nobr></td>
<%
PageFooter
end sub



sub SaveCorrespondence
  session("pagetab") = "People"
  session("history") = "Menu|p_menu.asp|Student Sets|p_sets.asp|Batch Processes|p_sets.asp?task=menu&id=" & id & "|Correspondence"
  set_query = "select 0 student_instance_id where 1=0"
  sql = "select set_query from student_sets where (set_query is not null) and student_set_id = " & checkstring(id,50)
  rs.open sql,conn,1,1
  if not rs.eof then
    set_query = rs("set_query")
  end if
  rs.close
  sql = "INSERT INTO correspondence (student_instance_id, letter_id, person_id, person_address_id, zone_name, status, requested_date, scheduled_date, printed_date, copies, notes, created_by, created_dt, updated_by, updated_dt) "
  sql = sql & " SELECT i.student_instance_id, "
  sql = sql & checkstring2(request("letter_id"),true) & " as letter_id, i.person_id, a.person_address_id, g.zone_name, "
  sql = sql & checkstring(request("status"),50) & ", " & checkstring(request("requested_date"),50) & ", "
  sql = sql & checkstring(request("scheduled_date"),50) & ", " & checkstring(request("printed_date"),50) & ", "
  sql = sql & checkstring(request("copies"),50) & ", " & checkstring2(request("notes"),false) & ", "
  sql = sql & checkstring(session("user_id"),50) & "," & checkstring(now,50) & ","
  sql = sql & checkstring(session("user_id"),50) & "," & checkstring(now,50)
  sql = sql & " FROM programs g, student_instance i LEFT OUTER JOIN person_address a"
  sql = sql & " ON i.person_id=a.person_id and a.address_type=" &  checkstring(request("addresstype"),30)
  sql = sql & " WHERE i.program_id = g.program_id and i.student_instance_id in (" & set_query & ") "
  conn.execute(sql)
  SaveComplete("The letter has been assigned to all students in this set.")
end sub





%>