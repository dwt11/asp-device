//���嵱ǰ�Ƿ��д��״̬
window.onload=
function() {
	password1=null;		
	initCalc();
	softkeyboard.style.display='none';
	}
var CapsLockValue=0;
var check;
function setVariables() {
	tablewidth=630;  // logo width, in pixels
	tableheight=20;  // logo height, in pixels
	if (navigator.appName == "Netscape") {
		horz=".left";
		vert=".top";
		docStyle="document.";
		styleDoc="";
		innerW="window.innerWidth";
		innerH="window.innerHeight";
		offsetX="window.pageXOffset";
		offsetY="window.pageYOffset";
		}
	else {
		horz=".pixelLeft";
		vert=".pixelTop";
		docStyle="";
		styleDoc=".style";
		innerW="document.body.clientWidth";
		innerH="document.body.clientHeight";
		offsetX="document.body.scrollLeft";
		offsetY="document.body.scrollTop";
   	}
}
function checkLocation() {
	if (check) {
		objectXY="softkeyboard";
		var availableX=eval(innerW);
		var availableY=eval(innerH);
		var currentX=eval(offsetX);
		var currentY=eval(offsetY);
		x=availableX-tablewidth+currentX;
		//y=availableY-tableheight+currentY;
		y=currentY;
		evalMove();
	}
	setTimeout("checkLocation()",0);
}
function evalMove() {
	//eval(docStyle + objectXY + styleDoc + horz + "=" + x);
	eval(docStyle + objectXY + styleDoc + vert + "=" + y);
}
self.onError=null;
currentX = currentY = 0;  
whichIt = null;           
lastScrollX = 0; lastScrollY = 0;
NS = (document.layers) ? 1 : 0;
IE = (document.all) ? 1: 0;
function heartBeat() {
	if(IE) { diffY = document.body.scrollTop; diffX = document.body.scrollLeft; }
	if(NS) { diffY = self.pageYOffset; diffX = self.pageXOffset; }
	if(diffY != lastScrollY) {
		percent = .1 * (diffY - lastScrollY);
		if(percent > 0) percent = Math.ceil(percent);
			else percent = Math.floor(percent);
		if(IE) document.all.softkeyboard.style.pixelTop += percent;
		if(NS) document.softkeyboard.top += percent; 
	        lastScrollY = lastScrollY + percent;
	}
	if(diffX != lastScrollX) {
		percent = .1 * (diffX - lastScrollX);
		if(percent > 0) percent = Math.ceil(percent);
		else percent = Math.floor(percent);
		if(IE) document.all.softkeyboard.style.pixelLeft += percent;
		if(NS) document.softkeyboard.left += percent;
		lastScrollX = lastScrollX + percent;	
	}
}
function checkFocus(x,y) { 
        stalkerx = document.softkeyboard.pageX;
        stalkery = document.softkeyboard.pageY;
        stalkerwidth = document.softkeyboard.clip.width;
        stalkerheight = document.softkeyboard.clip.height;
        if( (x > stalkerx && x < (stalkerx+stalkerwidth)) && (y > stalkery && y < (stalkery+stalkerheight))) return true;
	else return false;
}
function grabIt(e) {
	check = false;
	if(IE) {
		whichIt = event.srcElement;
		while (whichIt.id.indexOf("softkeyboard") == -1) {
			whichIt = whichIt.parentElement;
			if (whichIt == null) { return true; } }
				whichIt.style.pixelLeft = whichIt.offsetLeft;
				whichIt.style.pixelTop = whichIt.offsetTop;
				currentX = (event.clientX + document.body.scrollLeft);
	   			currentY = (event.clientY + document.body.scrollTop); 	
			} else { 
	        		window.captureEvents(Event.MOUSEMOVE);
	        		if(checkFocus (e.pageX,e.pageY)) { 
	                		whichIt = document.softkeyboard;
	                		StalkerTouchedX = e.pageX-document.softkeyboard.pageX;
	                		StalkerTouchedY = e.pageY-document.softkeyboard.pageY;
			}
		}
	return true;
}
function moveIt(e) {
	if (whichIt == null) { return false; }
		if(IE) {
			newX = (event.clientX + document.body.scrollLeft);
			newY = (event.clientY + document.body.scrollTop);
			distanceX = (newX - currentX);    distanceY = (newY - currentY);
			currentX = newX;    currentY = newY;
			whichIt.style.pixelLeft += distanceX;
			whichIt.style.pixelTop += distanceY;
			if(whichIt.style.pixelTop < document.body.scrollTop) whichIt.style.pixelTop = document.body.scrollTop;
			if(whichIt.style.pixelLeft < document.body.scrollLeft) whichIt.style.pixelLeft = document.body.scrollLeft;
			if(whichIt.style.pixelLeft > document.body.offsetWidth - document.body.scrollLeft - whichIt.style.pixelWidth - 20) whichIt.style.pixelLeft = document.body.offsetWidth - whichIt.style.pixelWidth - 20;
			if(whichIt.style.pixelTop > document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5) whichIt.style.pixelTop = document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5;
			event.returnValue = false;
		} else { 
			whichIt.moveTo(e.pageX-StalkerTouchedX,e.pageY-StalkerTouchedY);
	        	if(whichIt.left < 0+self.pageXOffset) whichIt.left = 0+self.pageXOffset;
	        	if(whichIt.top < 0+self.pageYOffset) whichIt.top = 0+self.pageYOffset;
	        	if( (whichIt.left + whichIt.clip.width) >= (window.innerWidth+self.pageXOffset-17)) whichIt.left = ((window.innerWidth+self.pageXOffset)-whichIt.clip.width)-17;
	        	if( (whichIt.top + whichIt.clip.height) >= (window.innerHeight+self.pageYOffset-17)) whichIt.top = ((window.innerHeight+self.pageYOffset)-whichIt.clip.height)-17;
	        	return false;
		}
	return false;
}

function dropIt() {
	whichIt = null;
	if(NS) window.releaseEvents (Event.MOUSEMOVE);
	return true;
}

if(NS) {window.captureEvents(Event.MOUSEUP|Event.MOUSEDOWN);
	window.onmousedown = grabIt;
	window.onmousemove = moveIt;
	window.onmouseup = dropIt;	}
if(IE) {
	document.onmousedown = grabIt;
 	document.onmousemove = moveIt;
	document.onmouseup = dropIt;	}

//	if(NS || IE) action = window.setInterval("heartBeat()",1);
document.write("<DIV align=center id=\"softkeyboard\" name=\"softkeyboard\" style=\"position:absolute; left:0px; top:0px; width:500px; z-index:180;display:none\"><table id=\"CalcTable\" width=\"\" border=\"0\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\" bgcolor=\"\"><FORM id=Calc name=Calc action=\"\" method=post autocomplete=\"off\"><tr><td align=\"center\" style=\"height:0\"><font color=#043BC9>==&nbsp;����������ʹ��������������������&nbsp;==</font></td></tr><tr align=\"center\"><td align=\"center\" bgcolor=\"#FFFFFF\"><INPUT type=hidden value=\"\" name=password><INPUT type=hidden value=ok name=action2><table align=\"center\" width=\"%\" border=\"0\" cellspacing=\"1\" cellpadding=\"0\">\n <tr align=\"left\" valign=\"middle\">\n <td><input type=button value=\" ~ \"></td>\n <td><input type=button value=\" ! \"></td>\n <td> <input type=button  value=\" @ \"></td>\n <td><input type=button value=\" # \"></td>\n <td><input type=button value=\" $ \"></td>\n <td><input type=button value=\" % \"></td>\n <td><input type=button value=\" ^ \"></td>\n <td><input type=button value=\" & \"></td>\n <td><input type=button value=\" * \"></td>\n <td><input type=button value=\" ( \"></td>\n <td><input type=button value=\" ) \"></td>\n <td><input type=button value=\" _ \"></td>\n <td><input type=button value=\" + \"></td>\n <td><input type=button value=\" | \"></td>\n <td colspan=\"1\" rowspan=\"2\"> <input name=\"button10\" type=button value=\"�˸�\" onclick=\"setpassvalue();\"  onDblClick=\"setpassvalue();\" style=\"width:40px;height:41px\">\n </td>\n </tr>\n <tr align=\"left\" valign=\"middle\"> \n <td><input type=button value=\" ` \"></td>\n <td><input type=button value=\" 1 \"></td>\n <td><input type=button value=\" 2 \"></td>\n <td><input type=button value=\" 3 \"></td>\n <td><input type=button value=\" 4 \"></td>\n <td><input type=button value=\" 5 \"></td>\n <td><input type=button value=\" 6 \"></td>\n <td><input type=button value=\" 7 \"></td>\n <td><input type=button value=\" 8 \"></td>\n <td><input type=button value=\" 9 \"></td>\n <td><input name=\"button6\" type=button value=\" 0 \"></td>\n <td><input type=button value=\" - \"></td>\n <td><input type=button value=\" = \"></td>\n <td><input type=button value=\" \\ \"></td>\n <td></td>\n </tr>\n <tr align=\"left\" valign=\"middle\">\n <td><input type=button value=\" q \"></td>\n <td><input type=button value=\" w \"></td>\n <td><input type=button value=\" e \"></td>\n <td><input type=button value=\" r \"></td>\n <td><input type=button value=\" t \"></td>\n <td><input type=button value=\" y \"></td>\n <td><input type=button value=\" u \"></td>\n <td><input type=button value=\" i \"></td>\n  <td><input type=button value=\" o \"></td>\n <td><input name=\"button8\" type=button value=\" p \"></td>\n <td><input name=\"button9\" type=button value=\" { \"></td>\n <td><input type=button value=\" } \"></td>\n <td><input type=button value=\" [ \"></td>\n <td><input type=button value=\" ] \"></td>\n <td><input name=\"capslock\" type=button onClick=\"setCapsLock();capsLockText();\" value=\"Сд\" style=\"width:40px;\"></td>\n </tr>\n <tr align=\"left\" valign=\"middle\">\n <td><input type=button value=\" a \"></td>\n <td><input type=button value=\" s \"></td>\n <td><input type=button value=\" d \"></td>\n <td><input type=button value=\" f \"></td>\n <td><input type=button value=\" g \"></td>\n <td><input type=button value=\" h \"></td>\n <td><input type=button value=\" j \"></td>\n <td><input name=\"button3\" type=button value=\" k \"></td>\n <td><input name=\"button4\" type=button value=\" l \"></td>\n <td><input name=\"button5\" type=button value=\" : \"></td>\n <td><input name=\"button7\" type=button value=\" &quot; \"></td>\n <td><input type=button value=\" ; \"></td>\n <td><input type=button value=\" ' \"></td>\n <td colspan=\"2\"><input name=\"button12\" type=button onclick=\"closekeyboard();\" value=\"��������\" style=\"width:66px\"></td>\n </tr>\n <tr align=\"left\" valign=\"middle\">\n <td><input name=\"button2\" type=button value=\" z \"></td>\n <td><input type=button value=\" x \"></td>\n <td><input type=button value=\" c \"></td>\n <td><input type=button value=\" v \"></td>\n <td><input type=button value=\" b \"></td>\n <td><input type=button value=\" n \"></td>\n <td><input type=button value=\" m \"></td>\n <td><input type=button value=\" &lt; \"></td>\n <td><input type=button value=\" &gt; \"></td>\n <td><input type=button value=\" ? \"></td>\n <td><input type=button value=\" , \"></td>\n <td><input type=button value=\" . \"></td>\n <td><input type=button value=\" / \"></td>\n <td colspan=\"2\"><input name=\"button12\" type=button onclick=\"OverInput();\" value=\"ȷ ��\" style=\"width:66px\"></td>\n </tr>\n </table></td></FORM></tr></table></DIV>")

//�����������������ֵ
function addValue(newValue) {
	if (CapsLockValue==0) {
		var str=Calc.password.value;
		if(str.length<password1.maxLength) {
			Calc.password.value += newValue;
		}			
		if(str.length<=password1.maxLength) {
			password1.value=Calc.password.value;
		}
	} else {
		var str=Calc.password.value;
		if(str.length<password1.maxLength) {
			Calc.password.value += newValue.toUpperCase();
		}
		if(str.length<=password1.maxLength) {
			password1.value=Calc.password.value;
		}
	}
}

//ʵ��BackSpace���Ĺ���
function setpassvalue()	{
	var longnum=Calc.password.value.length;
	var num
	num=Calc.password.value.substr(0,longnum-1);
	Calc.password.value=num;
	var str=Calc.password.value;
	password1.value=Calc.password.value;
}

//�������
function OverInput() {
	password1.value=Calc.password.value;
	softkeyboard.style.display="none";
	Calc.password.value="";
	password1.readOnly=1;
	document.Login.CheckCode.focus();
}

//�ر������
function closekeyboard() {
	Calc.password.value="";
	password1.readOnly=0;
	password1.value='';
	//softkeyboard.style.display='none';
	password1.focus();
	closestr=1;
}

//��ʾ�����
function showkeyboard() {
	if (softkeyboard.style.display == "none") {
	if(event.y+140)
	softkeyboard.style.top=event.y+document.body.scrollTop+15;
	if ((event.x-250)>0) {
		softkeyboard.style.left=event.x-250;
	} else {
		softkeyboard.style.left=0;
	}
	softkeyboard.style.display="block";
	password1.readOnly=1;
	password1.blur();
	// password1.value="";
	}
}

//�����Ƿ��д��ֵ
function setCapsLock() {
	if (CapsLockValue==0) {
		CapsLockValue=1;
		Calc.capslock.value="��д";
	} else {
		CapsLockValue=0;
		Calc.capslock.value="Сд";
	}
}

function setCalcborder() {
	CalcTable.style.border="1px solid #C4C2C4" //�߿���ɫ����
}

function setHead() {
	CalcTable.cells[0].style.backgroundColor="#ECEEEC"  //����������ɫ����
}

function setCalcButtonBg() {
	for(var i=0;i<Calc.elements.length;i++)	{
		if(Calc.elements[i].type=="button"&&Calc.elements[i].bgtype!="1") {
			//	if(i==10)
			//	alert(123);
			Calc.elements[i].style.borderTopWidth= 0
			Calc.elements[i].style.borderRightWidth= 2
			Calc.elements[i].style.borderBottomWidth= 2
			Calc.elements[i].style.borderLeftWidth= 0
			Calc.elements[i].style.borderTopStyle= "none";
			Calc.elements[i].style.borderRightStyle= "solid";
			Calc.elements[i].style.borderBottomStyle= "solid";
			Calc.elements[i].style.borderLeftStyle= "none";
			//#46AC17
			Calc.elements[i].style.borderTopColor= "#ECEEEC";��// ��ť��Ӱɫ���� ";
			Calc.elements[i].style.borderRightColor= "#ECEEEC";
			Calc.elements[i].style.borderBottomColor= "#ECEEEC";
			Calc.elements[i].style.borderLeftColor= "#ECEEEC";
			//#CBF3B2
			Calc.elements[i].style.backgroundColor="#F7F7F7"; //���̱���ɫ���� ";
			var str1=Calc.elements[i].value;
			str1=str1.trim();
			/*
			if(str1=="`") {
				Calc.elements[i].style.fontSize=14;
			}
			*/
			if(str1.length==1) {
				//Calc.elements[i].style.fontSize=16;
				//Calc.elements[i].style.fontWeight='bold';
			}
			var thisButtonValue=Calc.elements[i].value;
			thisButtonValue=thisButtonValue.trim();
			if(thisButtonValue.length==1) {
				Calc.elements[i].onclick=function () {
					var thisButtonValue=this.value;
					thisButtonValue=thisButtonValue.trim();
					addValue(thisButtonValue);
					//alert(234)
					}
				Calc.elements[i].ondblclick=function () {
					var thisButtonValue=this.value;
					thisButtonValue=thisButtonValue.trim();
					addValue(thisButtonValue);
					//alert(234)
				}
			}
		}
	}
}

function initCalc() {
	setCalcborder();
	setHead();
	setCalcButtonBg();
}
String.prototype.trim = function() {
    // ��������ʽ��ǰ��ո�
    // �ÿ��ַ��������
    return this.replace(/(^\s*)|(\s*$)/g, "");
}
var capsLockFlag;
capsLockFlag=true;

function capsLockText() {
	if(capsLockFlag) //�ĳɴ�д 
	{
		for(var i=0;i<Calc.elements.length;i++)	{
			var char=Calc.elements[i].value;
			var char=char.trim()
			if(Calc.elements[i].type=="button"&&char>="a"&&char<="z"&&char.length==1) {
				Calc.elements[i].value=" "+String.fromCharCode(char.charCodeAt(0)-32)+" "
			}
		}
	} else {
		for(var i=0;i<Calc.elements.length;i++) {
			var char=Calc.elements[i].value;
			var char=char.trim()
			if(Calc.elements[i].type=="button"&&char>="A"&&char<="Z"&&char.length==1) {
				Calc.elements[i].value=" "+String.fromCharCode(char.charCodeAt(0)+32)+" "
			}		}
	}
	capsLockFlag=!capsLockFlag;
}