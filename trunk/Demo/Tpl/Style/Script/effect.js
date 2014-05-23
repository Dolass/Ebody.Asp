///////////////////////////////////////////////////////////////////////////
// 图片倒影效果代码
///////////////////////////////////////////////////////////////////////////



///////////////////////////////////////////////////////////////////////////
// 方法一
///////////////////////////////////////////////////////////////////////////


function ShadowImg(p_elementId)
{
  
var odiv=document.getElementById(p_elementId);    
//var oimg=document.getElementsByTagName('img');
var oimg=odiv.getElementsByTagName('img');  
var img=oimg[0];    
if (document.createElement("canvas").getContext) 
	{    
		flx = document.createElement("canvas");    
		flx.width = img.width;  
		flx.height = img.height; 			
		var context = flx.getContext("2d");    
		context.translate(0, img.height);    
		context.scale(1, -1);    
		context.drawImage(img, 0, 0, img.width, img.height);    
		context.globalCompositeOperation = "destination-out";   
		
		// original script
		//var gradient = context.createLinearGradient(0, 0, 0, img.height * 3);  //设定Firefox的图片透明度 3  
		//gradient.addColorStop(1, "rgba(255, 255, 255, 0)");    
		//gradient.addColorStop(0, "rgba(255, 255, 255, 1)");

		// modify by tony 20091111
		var gradient = context.createLinearGradient(0, 50, 0, img.height);
		gradient.addColorStop(0, "rgba(255, 255, 255, 1)");    
		gradient.addColorStop(1, "rgba(255, 255, 255, 0.6)"); // 0.6 is opacity

		context.fillStyle = gradient;    
		context.fillRect(0, 0, img.width, img.height * 2);    
	} 
	else 
	{
		//ie浏览  
		var flx;  
		flx=document.createElement('img');   
		flx.src=img.src;   
		flx.style.filter='flipv progid:DXImageTransform.Microsoft.Alpha(' +    
                   'opacity=40, style=1, finishOpacity=0, startx=0, starty=0, finishx=0, finishy=' +    
                               (img.height * .25) + ')';  //设定IE的图片透明度 .25
		}  
  
flx.style.position = 'absolute';
flx.style.top = '301px';
flx.style.left = '0px';
//flx.style.float = 'right';
odiv.appendChild(flx);
  
}