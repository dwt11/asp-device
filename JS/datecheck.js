  function datecheck(tabn){      //判断日期(长日期格式为:2006-04-03 17:55:00)的正则
var dval1=tabn.value;
var r=/^(\d{0,4})-(0{0,1}[1-9]|1[0-2])-(0{0,1}[1-9]|[1-2]\d|3[0-1])$/;
if(!r.test(dval1)){
    alert('输入日期错误');
    tabn.focus();
    return false;}
else{
    var r1=/^0{0,4}$/;
    if(r1.test(RegExp.$1)){ 
           alert('年份不能为0');
             tabn.focus();
              return false;}
              var r2=/1[02]|0{0,1}[13578]/;' //小月
              if(!r2.test(RegExp.$2)){
                     if(parseInt(RegExp.$2)==2){
                            if(parseInt(RegExp.$1)%4==0){
                                 if(parseInt(RegExp.$3)>29){
                                       alert('闰年2月只有29天');
                                       tabn.focus();
                                       return false;}}
                             else{
                                 if(parseInt(RegExp.$3)>28){
                                       alert('2月只有28天');
                                       tabn.focus();
                                       return false;}}}
                             else{
                                if(parseInt(RegExp.$3)>30){
                                        alert('小月只有30天');
                                        tabn.focus();
                                        return false;}}}}
    }
