    function loadBar(fl)
//fl is show/hide flag
{
  var x,y;
  if (self.innerHeight)
  {// all except Explorer
    x = self.innerWidth;
    y = self.innerHeight;
  }
  else 
  if (document.documentElement && document.documentElement.clientHeight)
  {// Explorer 6 Strict Mode
   x = document.documentElement.clientWidth;
   y = document.documentElement.clientHeight;
  }
  else
  if (document.body)
  {// other Explorers
   x = document.body.clientWidth;
   y = document.body.clientHeight;
  }

    var el=document.getElementById('loader');
        if(null!=el)
        {
                var top = (y/2) - 50;
                var left = (x/2) - 150;
                if( left<=0 ) left = 10;
                el.style.position="absolute";
                el.style.visibility = (fl==1)?'visible':'hidden';
                el.style.display = (fl==1)?'block':'none';
                el.style.left = left + "px"
                el.style.top = top + "px";
                el.style.zIndex = 88;
        }
}