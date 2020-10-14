
////usermanagement.asp 注册用户通信班验证
var num=3;
var mon1=0;

//封装得到对像ID涵数
function getObj(objName){return(document.getElementById(objName));}

//input得到焦点效果
function showare(id){
        for(var i=1;i<=num;i++){
		if(i==id) {
		 getObj("sps"+i).className="reg2";
	}
		else{

		  if(i==num || i==(id-1)){
		     getObj("sps"+i).className="reg1";
			 }
		  else{
			 getObj("sps"+i).className="reg1";
			 }
		 }
	}
}

//过滤
function RTrim(str)
{
 var whitespace=new String(" \t\n\r");
 var s=new String(str);
 if (whitespace.indexOf(s.charAt(s.length-1))!=-1)
 {
  var i=s.length-1;
  while (i>=0 && whitespace.indexOf(s.charAt(i))!=-1)
  {
   i--;
  }
    s=s.substring(0,i+1);
  }
 return s;
}

//过滤
function LTrim(str)
{
 var whitespace=new String(" \t\n\r");
 var s=new String(str);
 if (whitespace.indexOf(s.charAt(0))!=-1)
 {
  var j=0, i = s.length;
  while (j<i && whitespace.indexOf(s.charAt(j))!=-1)
 {
  j++;
 }
   s=s.substring(j,i);
 }
  return s;
}

function Trim(str)
{
 return RTrim(LTrim(str));
}

//输入密码检测
//用户名检测
function myuser()
{
   mon1=0;
   allok();

 if (getObj("input1").value=="" || Trim(getObj("input1").value)=="")
 { 
   getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>请输入编号</font>";

//   mon1=0;
//   allok();
   return false;
 }else{	var qptjtz_bh=getObj("input1").value;
//	 alert qptjtz_bh;
	/*
	*--------------- GetResult() -----------------
	* GetResult()
	* 功能:通过XMLHTTP发送请求,返回结果.
	*--------------- GetResult() -----------------
	*/
	/* Create a new XMLHttpRequest object to talk to the Web server */
	var oBao = false;
	/*@cc_on @*/
	/*@if (@_jscript_version >= 5)
	try {
	  oBao = new ActiveXObject("Msxml2.XMLHTTP");
	} catch (e) {
	  try {
		oBao = new ActiveXObject("Microsoft.XMLHTTP");
	  } catch (e2) {
		oBao = false;
	  }
	}
	@end @*/
	if (!oBao && typeof XMLHttpRequest != 'undefined') {
	  oBao = new XMLHttpRequest();
	}
	//特殊字符：+,%,&,=,?等的传输解决办法.字符串先用escape编码的.
	//Update:2004-6-1 12:22
		var userInfoo = "qptjtz_bh="+qptjtz_bh;
		oBao.open("POST","checkbh.asp",false);
		oBao.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
		oBao.send(userInfoo);
	//服务器端处理返回的是经过escape编码的字符串.
	var strResult =unescape(oBao.responseText);
	
	
		oBao.open("POST","checkbh.asp",false);
		oBao.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
		oBao.send(userInfoo);
	//服务器端处理返回的是经过escape编码的字符串.
	var stiResult =unescape(oBao.responseText);

	//alert (strResult);
	//将字符串分开.
	if (strResult==getObj("input1").value.toLowerCase()){
	getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><a href=qptjtz_whjl.asp?action=add&qptjtzid="+stiResult+">气瓶已存在，点击添加</a>";
//	mon1=0;
//	allok();
	return false;
	}else{
	getObj("sps1").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>继续添加！</font>";
//	mon1=1;
//	allok();
	return false;
	}
 }
}

//发送扭钮状态检测
function allok(){
  if (mon1==1){
getObj("submit").disabled=false;
}else{
getObj("submit").disabled=true;
}
}

  
 function  DateAdd(interval,number,date)
{
/*
  *   功能:实现VBScript的DateAdd功能.
  *   参数:interval,字符串表达式，表示要添加的时间间隔.
  *   参数:number,数值表达式，表示要添加的时间间隔的个数.
  *   参数:date,时间对象.
  *   返回:新的时间对象.
  *   var   now   =   new   Date();
  *   var   newDate   =   DateAdd( "d ",5,now);
  *---------------   DateAdd(interval,number,date)   -----------------
  */
  number = parseInt(number);  
if (typeof(date)=="string"){  
date = date.split(/\D/);  
--date[1];  
eval("var date = new Date("+date.join(",")+")");  
}  
if (typeof(date)=="object"){  
var date = date  
}  
switch(interval){  
case "y": date.setFullYear(date.getFullYear()+number); break;  
case "m": date.setMonth(date.getMonth()+number); break;  
case "d": date.setDate(date.getDate()+number); break;  
case "w": date.setDate(date.getDate()+7*number); break;  
case "h": date.setHours(date.getHour()+number); break;  
case "n": date.setMinutes(date.getMinutes()+number); break;  
case "s": date.setSeconds(date.getSeconds()+number); break;  
case "l": date.setMilliseconds(date.getMilliseconds()+number); break;  
}   
return date;  
} 

function addrdata(){
	mon1=1;
	allok();

    var yxq=document.getElementById("qptjtz_yxq").value;
	var lydata=document.getElementById("qptjtz_scdata").value;
	var dqdata1=DateAdd("m",yxq,lydata);
	var dqdata=dqdata1.format('yy-MM-dd');

	getObj("input6").value=dqdata;
}

