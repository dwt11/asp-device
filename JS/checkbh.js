
////usermanagement.asp ע���û�ͨ�Ű���֤
var num=3;
var mon1=0;

//��װ�õ�����ID����
function getObj(objName){return(document.getElementById(objName));}

//input�õ�����Ч��
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

//����
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

//����
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

//����������
//�û������
function myuser()
{
   mon1=0;
   allok();

 if (getObj("input1").value=="" || Trim(getObj("input1").value)=="")
 { 
   getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>��������</font>";

//   mon1=0;
//   allok();
   return false;
 }else{	var qptjtz_bh=getObj("input1").value;
//	 alert qptjtz_bh;
	/*
	*--------------- GetResult() -----------------
	* GetResult()
	* ����:ͨ��XMLHTTP��������,���ؽ��.
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
	//�����ַ���+,%,&,=,?�ȵĴ������취.�ַ�������escape�����.
	//Update:2004-6-1 12:22
		var userInfoo = "qptjtz_bh="+qptjtz_bh;
		oBao.open("POST","checkbh.asp",false);
		oBao.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
		oBao.send(userInfoo);
	//�������˴����ص��Ǿ���escape������ַ���.
	var strResult =unescape(oBao.responseText);
	
	
		oBao.open("POST","checkbh.asp",false);
		oBao.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
		oBao.send(userInfoo);
	//�������˴����ص��Ǿ���escape������ַ���.
	var stiResult =unescape(oBao.responseText);

	//alert (strResult);
	//���ַ����ֿ�.
	if (strResult==getObj("input1").value.toLowerCase()){
	getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><a href=qptjtz_whjl.asp?action=add&qptjtzid="+stiResult+">��ƿ�Ѵ��ڣ�������</a>";
//	mon1=0;
//	allok();
	return false;
	}else{
	getObj("sps1").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>������ӣ�</font>";
//	mon1=1;
//	allok();
	return false;
	}
 }
}

//����Ťť״̬���
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
  *   ����:ʵ��VBScript��DateAdd����.
  *   ����:interval,�ַ������ʽ����ʾҪ��ӵ�ʱ����.
  *   ����:number,��ֵ���ʽ����ʾҪ��ӵ�ʱ�����ĸ���.
  *   ����:date,ʱ�����.
  *   ����:�µ�ʱ�����.
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

