var BoxWidth = 480	// 资料表显示宽度 ( 不含卷轴 )
var ShowLine = 10	// 资料表显示列数
var RsHeight = 25	// 资料列高度
var LockCols = 2	// 要锁定的栏位数 ( 由左至右 )

function WriteTable(){	// 写入表格
var iBoxWidth=BoxWidth
var NewHTML="<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tr>\
<td><div style=\"width:100%;overflow-x:scroll\">\
<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tr>"
for(i=0;i<DataTitles.length;i++){
  if(i<LockCols){
    var cTitle=DataTitles[i].split("#")
    iBoxWidth-=cTitle[1]
    var DynTip=((i+1)==LockCols)?"解除锁定":"锁定此栏位"
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
    NewHTML+="<td><div class=\"title\" style=\"width:"+cTitle[1]+"px;height:"+RsHeight+"px\" title=\"锁定此栏位\" onclick=\"ResetTable("+i+")\">"+cTitle[0]+"</div></td>"
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

function ApplyData(){	// 写入资料
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
  var Sure=confirm("\n锁定栏位的宽度大於资料表显示的宽　　\n\n度，这可能会造成版面显示不正常。\n\n\n您确定要继续吗？")
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
