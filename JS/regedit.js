
////usermanagement.asp 注册用户通信班验证
var num=3;
var mon1=0;
var mon2=0;
var mon3=0;
var mon4=0;
var mon5=0;

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
function checkpassword()
{ 

 if (getObj("input2").value=="" || Trim(getObj("input2").value)=="")
 { 
   getObj("sps2").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>请输入注册密码！</font>";

   mon2=0;
   allok();
   return false;
 }else{
 if (Trim(getObj("input2").value).indexOf(" ")>=0) 
 { 
   getObj("sps2").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>密码中不能包含空格！</font>";

   mon2=0;
   allok();
  return false; 
 }else{
 if (Trim(getObj("input2").value).length<4) 
 { 
   getObj("sps2").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>密码不能少于4个字符！</font>";

   mon2=0;
   allok();
  return false; 
 }else{
 if (Trim(getObj("input2").value).length>20) 
 { 
   getObj("sps2").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>密码不能超过20个字符！</font>";

   mon2=0;
   allok();
  return false;
 }else{
   getObj("sps2").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>密码设置OK啦！</font>"; 
   mon2=1;
   allok(); 
  return false;
}
 }
 }
 }
}

//重复输入密码检测
function checkreturnpass()
{ 

 if (getObj("input3").value=="" || Trim(getObj("input3").value)=="")
 { 
   getObj("sps3").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>请确认密码！</font>";

   mon3=0;
   allok();
   return false;
 }else{
   if(getObj("input2").value!=getObj("input3").value)
   {
     getObj("sps3").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>您两次输入的密码不相符！</font>"

   mon3=0;
   allok();
     return false;
   }else{
    getObj("sps3").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>密码设置正确！</font>"
   mon3=1;
   allok();   
  return false; 
}
}
}

//权限检查
function checklevelclass(){
  if (getObj("input5").value=="" || Trim(getObj("input5").value)=="")
 { 
   getObj("sps5").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>请选择权限！</font>";

   mon5=0;
   allok();
   return false;
 }else{
    getObj("sps5").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>权限已选择！</font>"
   mon5=1;
   allok();   
  return false; 
}
}

////名字检查
//function checkusername1(){
//  if (getObj("input4").value=="" || Trim(getObj("input4").value)=="")
// { 
//   getObj("sps4").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>请添写！</font>";
//
//   mon4=0;
//   allok();
//   return false;
// }else{
//    getObj("sps4").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>权限已选择！</font>"
//   mon4=1;
//   allok();   
//  return false; 
//}
//}


//用户名检测
function myuser()
{
 if (getObj("input1").value=="" || Trim(getObj("input1").value)=="")
 { 
   getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>请输入用户名！</font>";

   mon1=0;
   allok();
   return false;
 }else{
 if (Trim(getObj("input1").value).indexOf(" ")>=0) 
 { 
   getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>用户名中不能包含空格！</font>";

   mon1=0;
   allok();
  return false; 
 }else{
 if (Trim(getObj("input1").value).length<3) 
 { 
   getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>用户名不能少于3个字符！</font>";

   mon1=0;
   allok();
  return false; 
 }else{	var username=getObj("input1").value;
	 //alert username;
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
		var userInfoo = "username="+username;
		oBao.open("POST","checkuser.asp",false);
		oBao.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
		oBao.send(userInfoo);
	//服务器端处理返回的是经过escape编码的字符串.
	var strResult = unescape(oBao.responseText);
	//alert (strResult);
	//将字符串分开.
	if (strResult==getObj("input1").value.toLowerCase()){
	getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>不行啦，有重名啦Q！</font>";
	mon1=0;
	allok();
	return false;
	}else{
	getObj("sps1").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>用户名OK啦！可以进行下步啦！</font>";
	mon1=1;
	allok();
	return false;
	}
 }
 }
 }
}

//发送扭钮状态检测
function allok(){
  if (mon1==1&&mon2==1&&mon3==1&&mon5==1){
getObj("submit").disabled=false;
}else{
getObj("submit").disabled=true;
}
}
