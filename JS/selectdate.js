function YMD(){ 
    YMD.year=document.getElementsByName(arguments[0])[0];
    YMD.month=document.getElementsByName(arguments[1])[0];
    YMD.day=document.getElementsByName(arguments[2])[0];

    this.year=document.getElementsByName(arguments[0])[0];
    this.month=document.getElementsByName(arguments[1])[0];
    this.day=document.getElementsByName(arguments[2])[0];
    $year=arguments[3].split("-");
    this.date=arguments[4].split("-");
    for(x=$year[0];x<=$year[1];x++){
        option=document.createElement("option");
        option.text=x
        option.value=x
        this.year.add(option);
        if(x==this.date[0]){
            option.selected=1
        }
    }
    this.setDay=function(obj){
        obj.day.length=0
        if(",1,3,5,7,8,10,12".indexOf(","+obj.date[1])!=-1){
            days=31
        }else{
            days=30
        }
        if(obj.date[1]==2){
            obj.date[0]=parseInt(obj.date[0]);
            if(!(obj.date[0]%4)){
                if(!(obj.date[0]%100)){
                    days=28;
                    if(!(obj.date[0]%400)) days=29;
                }else{
                    days=29;
                }
            }else{
                days=28;
            }
        }
        
        for(y=1;y<=days;y++){
            option=document.createElement("option");
            option.text=y
            option.value=y
            obj.day.add(option);
            if(y==obj.date[2]){
                option.selected=1
            }
        }
    }
    for(x=1;x<=12;x++){
        option=document.createElement("option");
        option.text=x
        option.value=x
        this.month.add(option);
        if(x==this.date[1]){
            option.selected=1
        }
       
    }
    this.setDay(this);
    setDay=function(){
        YMD.day.length=0
        if(",1,3,5,7,8,10,12".indexOf(","+YMD.month.value)!=-1){
            days=31
        }else{
            days=30
        }
        if(YMD.month.value==2){
            if(!(YMD.year.value%4)){
                if(!(YMD.year.value%100)){
                    days=28;
                    if(!(YMD.year.value%400)) days=29;
                }else{
                    days=29;
                }
            }else{
                days=28;
            }
        }
        for(y=1;y<=days;y++){
            option=document.createElement("option");
            option.text=y
            option.value=y
            YMD.day.add(option);
        }
    }
    this.month.attachEvent("onchange",setDay);
    this.year.attachEvent("onchange",setDay);
}
//var d = new Date()
//var vYear = d.getFullYear()
var vYear = this.form.getyear.value
//alert(this.form.getyear11.value);
var vMon = this.form.getmonth.value
var vDay = this.form.getday.value

new YMD("year","month","day","2008-2015",vYear+"-"+vMon+"-"+vDay);
function towwsscj(sscjno){
	window.location.href=sscjno;
}
