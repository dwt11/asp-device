var oPopup = window.createPopup();
var popTop=50;
function popmsg(msgstr){
var winstr="<table width=\"241\" height=\"172\" border=\"1\" cellpadding=\"0\" cellspacing=\"0\" style=\"border: 1 solid #0000CC\">";
winstr+="<tr><td height=\"23\"><div align=\"center\">邮件消息提示<\/div><\/td>";
winstr+="<\/tr><tr ><td style=\"font-size:12px; color: red; face: Tahoma\">"+msgstr+"<\/td><\/tr><\/table>";
oPopup.document.body.innerHTML = winstr;
popshow();
}
function popshow(){
window.status=popTop;
if(popTop>1720){
clearTimeout(mytime);
oPopup.hide();
return;
}else if(popTop>1520&&popTop<1720){
oPopup.show(screen.width-250,screen.height,241,1720-popTop);
}else if(popTop>1500&&popTop<1520){
oPopup.show(screen.width-250,screen.height+(popTop-1720),241,172);
}else if(popTop<180){
oPopup.show(screen.width-250,screen.height,241,popTop);
}else if(popTop<220){
oPopup.show(screen.width-250,screen.height-popTop,241,172);
}
popTop+=10;
var mytime=setTimeout("popshow();",50);
}
