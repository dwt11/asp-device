//ȥ��URL�е�sscj,URL�д���
function  tosscj(sscjno){
	var dwturl=document.URL;
	dwturl=dwturl+'&sscj='+sscjno; //��JS�в��������У���ֹ��һҳתĬ����ʾʱû��PAGE����
	dwturl=dwturl.replace(/&sscj\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&page\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&keyword\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&pgsz\=(\d*)/g,'');  
	dwturl=dwturl.replace(/&ssgh\=(\d*)/g,'');  
	dwturl=dwturl+'&sscj='+sscjno;
	window.location.href=dwturl;
}

//ȥ��URL�е�ssgh,URL�д���
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


