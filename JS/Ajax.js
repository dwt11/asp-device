var navigatorname;
if (navigator.userAgent.indexOf("Opera")>=0)
{
  navigatorname="Opera"
}
if (navigator.userAgent.indexOf("MSIE")>=0)
{ 
navigatorname="IE"
}
if (navigator.userAgent.indexOf("Firefox")>=0)
{ 
navigatorname="Firefox"
}

//AJAX 对象
var XmlHTTP_obj=function(){
	if (window.XMLHttpRequest) { // Mozilla, Safari, ...
    return new XMLHttpRequest();
} else if (window.ActiveXObject) { // IE
    return new ActiveXObject("Microsoft.XMLHTTP");
}
}

//用于获取相应地址返回的XML档
function GetXMLContent(urlstr){
	var http_request;
	urlstr.indexOf("?")==-1?urlstr=urlstr+"?"+Math.random():urlstr=urlstr+"&"+Math.random()
  if (window.XMLHttpRequest){ // Mozilla, Safari, ...
    http_request = new XMLHttpRequest();
    } 
    else if (window.ActiveXObject) { // IE
    http_request = new ActiveXObject("Microsoft.XMLHTTP");
    }
    http_request.open("GET",urlstr,false);
    http_request.send(null);
    return http_request.responseXML;
	}

//使FIREFOX支持selectNodes()、selectSingleNode()
//代码出处：http://km0ti0n.blunted.co.uk/mozXPath.xap
// check for XPath implementation
if( document.implementation.hasFeature("XPath", "3.0") )
{
// prototying the XMLDocument
XMLDocument.prototype.selectNodes = function(cXPathString, xNode)
{
if( !xNode ) { xNode = this; } 
var oNSResolver = this.createNSResolver(this.documentElement)
var aItems = this.evaluate(cXPathString, xNode, oNSResolver, 
XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)
var aResult = [];
for( var i = 0; i < aItems.snapshotLength; i++)
{
aResult[i] = aItems.snapshotItem(i);
}
return aResult;
}

// prototying the Element
Element.prototype.selectNodes = function(cXPathString)
{
if(this.ownerDocument.selectNodes)
{
  return this.ownerDocument.selectNodes(cXPathString, this);
}
else{throw "For XML Elements Only";}
}
}

// check for XPath implementation
if( document.implementation.hasFeature("XPath", "3.0") )
{
// prototying the XMLDocument
XMLDocument.prototype.selectSingleNode = function(cXPathString, xNode)
{
if( !xNode ) { xNode = this; } 
var xItems = this.selectNodes(cXPathString, xNode);
if( xItems.length > 0 )
{
return xItems[0];
}
else
{
return null;
}
}

// prototying the Element
Element.prototype.selectSingleNode = function(cXPathString)
{ 
if(this.ownerDocument.selectSingleNode)
{
return this.ownerDocument.selectSingleNode(cXPathString, this);
}
else{throw "For XML Elements Only";}
}
}