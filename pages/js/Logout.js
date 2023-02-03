function MyLink(LinkFile)
{
//	var addr = "https://" + location.hostname + ":" + location.port +"/" + LinkFile;
	var addr = "http://" + location.hostname +"/pages/" + LinkFile;
	parent.location = addr;
}