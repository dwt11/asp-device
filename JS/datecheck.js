  function datecheck(tabn){      //�ж�����(�����ڸ�ʽΪ:2006-04-03 17:55:00)������
var dval1=tabn.value;
var r=/^(\d{0,4})-(0{0,1}[1-9]|1[0-2])-(0{0,1}[1-9]|[1-2]\d|3[0-1])$/;
if(!r.test(dval1)){
    alert('�������ڴ���');
    tabn.focus();
    return false;}
else{
    var r1=/^0{0,4}$/;
    if(r1.test(RegExp.$1)){ 
           alert('��ݲ���Ϊ0');
             tabn.focus();
              return false;}
              var r2=/1[02]|0{0,1}[13578]/;' //С��
              if(!r2.test(RegExp.$2)){
                     if(parseInt(RegExp.$2)==2){
                            if(parseInt(RegExp.$1)%4==0){
                                 if(parseInt(RegExp.$3)>29){
                                       alert('����2��ֻ��29��');
                                       tabn.focus();
                                       return false;}}
                             else{
                                 if(parseInt(RegExp.$3)>28){
                                       alert('2��ֻ��28��');
                                       tabn.focus();
                                       return false;}}}
                             else{
                                if(parseInt(RegExp.$3)>30){
                                        alert('С��ֻ��30��');
                                        tabn.focus();
                                        return false;}}}}
    }
