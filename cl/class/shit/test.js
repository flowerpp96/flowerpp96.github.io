function CA()
{
  for (var i=0;i<document.test.elements.length;i++)
    {
    var e=document.test.elements[i];
    if (e.name!='allbox')
      e.checked=document.test.allbox.checked;
    }
}

function doSection(ss)
{	
	img=eval("document.test."+ss.id+"_i")
	if (ss.style.visibility=="visible")
	{
		ss.style.visibility="hidden"
		img.src="image/right.gif"
		img.height="16"
		img.width="9"
	}
	else
	{
		ss.style.visibility="visible"
		img.src="image/down_red.gif"
		img.height="8"
		img.width="15"
	}
}