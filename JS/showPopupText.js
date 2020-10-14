var oPopup = window.createPopup();
var oPopupBody = oPopup.document.body;
function pop(msg,bgcolor)
{
	oPopupBody.style.fontSize = "9pt";
	oPopupBody.bgColor = bgcolor;
	oPopupBody.style.color = "#FFFFcc";
	oPopupBody.style.padding = "10";
  oPopupBody.innerHTML = msg;
  oPopup.show(0, 20, 0, 104, window.event.srcElement);
	var PopWinWidth = oPopupBody.scrollWidth;
	if(PopWinWidth<150)PopWinWidth=200;
	oPopup.hide();
  oPopup.show(-(PopWinWidth-window.event.srcElement.scrollWidth)/2, 20, PopWinWidth, 104, window.event.srcElement);
}
function kill()
{
	oPopup.hide();
}
