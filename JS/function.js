//去除URL中的sscj,URL中带？
function  tosscj(sscjno){
	var dwturl=document.URL;
	dwturl=dwturl+'&sscj='+sscjno; //此JS中参数必须有，防止第一页转默认显示时没有PAGE参数
	dwturl=dwturl.replace(/&sscj\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&page\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&keyword\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&pgsz\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&ssgh\=(\d*)/g,'');  
	dwturl=dwturl+'&sscj='+sscjno;
	window.location.href=dwturl;
}

//去除URL中的ssgh,URL中带？
function  tossgh(ssghno){
	var dwturl=document.URL;
	dwturl=dwturl+'&ssgh='+ssghno;
	dwturl=dwturl.replace(/&ssgh\=(\d*)/g,'');
	dwturl=dwturl.replace(/&sscj\=(\d*)/g,'');
	dwturl=dwturl.replace(/&keyword\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&page\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&pgsz\=(\d*)/g,'');  
	dwturl=dwturl+'&ssgh='+ssghno;
	window.location.href=dwturl;
}


