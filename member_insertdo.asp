<%
  Set fs = Server.CreateObject("Scripting.FileSystemObject")
  Set objFile = fs.OpenTextFile("C:\inetpub\wwwroot\web\test.txt", 8, true)

  member_number = request("member_number")
  member_name = request("member_name")
  memtype_no = request("memtype_no")
  sex = request("sex")
  hand_tel = request("hand_tel")
  home_tel = request("home_tel")
  home_addr = request("home_addr")
  e_mail = request("e_mail")
  deposit = request("deposit")
  join_day = request("join_day")
  why_join = request("why_join")

  objFile.writeLine(member_number)
  objFile.writeLine(member_name)
  objFile.writeLine(memtype_no)
  objFile.writeLine(sex)
  objFile.writeLine(hand_tel)
  objFile.writeLine(home_tel)
  objFile.writeLine(home_addr)
  objFile.writeLine(e_mail)
  objFile.writeLine(deposit)
  objFile.writeLine(join_day)
  objFile.writeLine(why_join)
  objFile.writeLine()

  objFile.close
%>

<HTML>
<BODY>
<br><center><font face="돋움" size="2">
<p style="color:red; font-size:200px;">회원가입</p>
<p style="color:red; font-size:90px;">회원 가입이 완료되었습니다.</p>
</font></center></BODY>
</HTML>




