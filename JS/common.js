//��ȡID
function $id(id){
	return document.getElementById(id);
	}
	
//��ȡtagname
function $tag(tanme){
	return document.getElementsByTagName(tanme);
	}
//��һ��
function disply(dclassid){
//alert(dclassid);
	//var iddiv_obj=$id("Src_ID"+srcid);
	var classdiv_obj="sb_dclass";
	var RetCode,RetDes;
  var xmldocumento=GetXMLContent("js/doajax.asp?action=disply&sd_dclassid="+dclassid);
  	alert("js/doajax.asp?action=disply&sd_dclassid="+dclassid);
  RetCode=xmldocumento.selectSingleNode( "//ReturnStr/RetCode/text()").nodeValue;
  RetDes=xmldocumento.selectSingleNode( "//ReturnStr/RetDes/text()").nodeValue;
  switch(RetCode){
  	case "0000" :
  	  classdiv_obj.innerHTML=(11111).toString();
  	  //iddiv_obj.innerHTML="�Ѷ�";
  	  break;

	case "0001" :
  	  alert(RetDes);
  	   break;
  	case "0002" :
  	  hitdiv_obj.innerHTML=(Math.round(hitdiv_obj.innerHTML)+1).toString();
  	  iddiv_obj.innerHTML="�Ѷ�";
  	  alert(RetDes);
  	   break;
  	case "0003" :
  	  hitdiv_obj.innerHTML=(Math.round(hitdiv_obj.innerHTML)+1).toString();
  	  iddiv_obj.innerHTML="�Ѷ�";
  	  alert(RetDes);
  	  break;
  	}
  	//window.status=Web_StatusKey;
	}

//�����¼�
function CtreateEvent(obj,eventname,func){
	if(navigator.userAgent.indexOf("MSIE")>=0){
		var f =new Function("event",func);
		obj.attachEvent(eventname,f);
		}
		else{
			obj.setAttribute(eventname,func);
			}
	}
//���ڻ��һ��ҳ���RSS����
	function getRssUrl(){
		var theurl="";
		theurl=document.getElementsByTagName("link")[3].href
		return theurl;
		}