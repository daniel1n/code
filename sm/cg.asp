         <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><p>          <%
on error resume next
'�����û���Ϣ
'��
dim sqlnian,weight1
sqlnian="select * from chenggu1 where year='"&nian1&"'"
set rsnian=conn.execute(sqlnian)
if not (rsnian.bof and rsnian.eof) then
weight1=rsnian("weight")
end if
'��
dim sqlyue,weight2
sqlyue="select * from chenggu2 where month='"&yue2&"'"
set rsyue=conn.execute(sqlyue)
if not (rsyue.bof and rsyue.eof) then
weight2=rsyue("weight")
end if
'��
dim sqlri,weight3
sqlri="select * from chenggu3 where day='"&ri2&"'"
set rsri=conn.execute(sqlri)
if not (rsri.bof and rsri.eof) then
weight3=rsri("weight")
end if
'ʱ��
dim sqlhh,weight4
sqlhh="select * from chenggu4 where time='"&hh2&"'"
set rshh=conn.execute(sqlhh)
if not (rshh.bof and rshh.eof) then
weight4=rshh("weight")
end if
'����������
dim weight
weight=weight1+weight2+weight3+weight4
%>
              <%
'�ƹ�����
dim tempweight,sqlchenggu,chenggucontent
tempweight=cstr(int(weight*10))
sqlchenggu="select * from chenggu where weight='"&tempweight&"'"
set rschenggu=conn.execute(sqlchenggu)
if not (rschenggu.bof and rschenggu.eof) then
chenggucontent=rschenggu("content")
intro=rschenggu("intro")

end if 
%>
      <%
'�����û���Ϣ
userbirthday=nian1&"-"&yue1&"-"&ri1
userll=DateDiff("yyyy", userbirthday, date)
if userll<=12 then
cf="�ɰ��ĵ�"
elseif userll>12 and userll<=50 then
cf="�װ���"
elseif userll>50 then
cf="�𾴵�"
end if
%>    Ԭ��ƹ�����</p>
          <p><%=cf%><span style="PADDING-BOTTOM: 1px"><%=xing&ming%></span>,����������������,�������㣬���Ĺ���Ϊ��<span class="red"><%=weight%></span> �� ��������(�����ο�)��</p>
          <hr>
              <div align="center" style="font-size:16px; font-weight:bold; color:#FF0000; line-height:40px">
                <%=chenggucontent%>
              </div>
              <hr>
                ������ͣ�<br>
                <br> <font color="#0000FF">
                <%=intro%></font><br>
<div id="ft"><%=now()%> by Senlon</div></TD>
      </tr>
    </TBODY>
  </TABLE>
