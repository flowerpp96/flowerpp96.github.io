
function doSection(ss)
{	
	img=eval("document."+ss.id+"_i")
	if (ss.style.visibility=="visible")
	{
		ss.style.visibility="hidden"
		img.src="images/up.gif"
		img.height="8"
		img.width="14"
	}
	else
	{
		ss.style.visibility="visible"
		img.src="images/right.gif"
		img.height="16"
		img.width="9"
	}
}
function doSection1(ss)
{	
	img=eval("document."+ss.id+"_i")
	if (ss.style.visibility=="visible")
	{
		img.src="images/down_red.gif"
		img.height="8"
		img.width="15"
		ss.style.visibility="hidden"

	}
	else
	{
		img.src="images/right.gif"
		img.height="16"
		img.width="9"
		ss.style.visibility="visible"
		
	}
}