var BoxWidth = 480	// ���ϱ���ʾ��� ( �������� )
var ShowLine = 10	// ���ϱ���ʾ����
var RsHeight = 25	// �����и߶�
var LockCols = 2	// Ҫ��������λ�� ( �������� )

function WriteTable(){	// д����
var iBoxWidth=BoxWidth
var NewHTML="<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tr>\
<td><div style=\"width:100%;overflow-x:scroll\">\
<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tr>"
for(i=0;i<DataTitles.length;i++){
  if(i<LockCols){
    var cTitle=DataTitles[i].split("#")
    iBoxWidth-=cTitle[1]
    var DynTip=((i+1)==LockCols)?"�������":"��������λ"
    NewHTML+="<td><div class=\"title\" style=\"width:"+cTitle[1]+"px;height:"+RsHeight+"px\" title=\""+DynTip+"\" onclick=\"ResetTable("+i+")\">"+cTitle[0]+"</div></td>"
  }
}
NewHTML+="</tr>\
<tr><td colspan=\""+LockCols+"\">\
<div id=\"DataFrame1\" style=\"position:relative;width:100%;overflow:hidden\">\
<div id=\"DataGroup1\" style=\"position:relative\"></div></div>\
</td></tr></table></div></td>\
<td valign=\"top\"><div style=\"width:"+iBoxWidth+"px;overflow-x:scroll\">\
<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tr>"
for(i=0;i<DataTitles.length;i++){
  if(i>=LockCols){
    var cTitle=DataTitles[i].split("#")
    NewHTML+="<td><div class=\"title\" style=\"width:"+cTitle[1]+"px;height:"+RsHeight+"px\" title=\"��������λ\" onclick=\"ResetTable("+i+")\">"+cTitle[0]+"</div></td>"
  }
}
NewHTML+="</tr>\
<tr><td colspan=\""+(DataTitles.length-LockCols)+"\">\
<div id=\"DataFrame2\" style=\"position:relative;width:100%;overflow:hidden\">\
<div id=\"DataGroup2\" style=\"position:relative\"></div>\
</div></td></tr></table>\
</div></td><td valign=\"top\">\
<div id=\"DataFrame3\" style=\"position:relative;background:#000;overflow-y:scroll\" onscroll=\"SYNC_Roll()\">\
<div id=\"DataGroup3\" style=\"position:relative;width:1px;visibility:hidden\"></div>\
</div></td></tr></table>"
DataTable.innerHTML=NewHTML
ApplyData()
}

function ApplyData(){	// д������
var NewHTML="<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\">"
for(i=0;i<DataFields.length;i++){
  NewHTML+="<tr>"
  for(j=0;j<DataTitles.length;j++){
    if(j<LockCols){
      var cTitle=DataTitles[j].split("#")
      NewHTML+="<td><div class=\"cdata\" style=\"width:"+cTitle[1]+"px;height:"+RsHeight+"px;text-align:"+cTitle[2]+"\">"+DataFields[i][j]+"</div></td>"
    }
  }
  NewHTML+="</tr>"
}
NewHTML+="</table>"
DataGroup1.innerHTML=NewHTML


var NewHTML="<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\">"
for(i=0;i<DataFields.length;i++){
  NewHTML+="<tr>"
  for(j=0;j<DataTitles.length;j++){
    if(j>=LockCols){
      var cTitle=DataTitles[j].split("#")
      
	  NewHTML+="<td><div class=\"cdata\" style=\"width:"+cTitle[1]+"px;height:"+RsHeight+"px;text-align:"+cTitle[2]+"\">"+DataFields[i][j]+"</div></td>"
    }
  }
  NewHTML+="</tr>"
}
NewHTML+="</table>"
DataGroup2.innerHTML=NewHTML
DataFrame1.style.pixelHeight=RsHeight*ShowLine
DataFrame2.style.pixelHeight=RsHeight*ShowLine
DataFrame3.style.pixelHeight=RsHeight*ShowLine+RsHeight
DataGroup3.style.pixelHeight=RsHeight*(DataFields.length+1)
}

function ResetTable(n){
var iBoxWidth=0
for(i=0;i<DataTitles.length;i++){
  if(i<(n+1)){
    var cTitle=DataTitles[i].split("#")
    iBoxWidth+=parseInt(cTitle[1])
  }
}
if(iBoxWidth>BoxWidth){
  var Sure=confirm("\n������λ�Ŀ�ȴ�����ϱ���ʾ�Ŀ���\n\n�ȣ�����ܻ���ɰ�����ʾ��������\n\n\n��ȷ��Ҫ������")
}else{
  Sure=true
}
if(Sure){
  LockCols=(LockCols==n+1)?0:n+1
  WriteTable()
}
}

function SYNC_Roll(){
DataGroup1.style.posTop=-DataFrame3.scrollTop
DataGroup2.style.posTop=-DataFrame3.scrollTop
}
window.onload=WriteTable
