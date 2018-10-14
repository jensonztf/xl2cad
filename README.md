# xl2cad
put text in to autocad as dbtext from wps.et or ms.excel,while get them out on the otherhand.</br>
the dll build in x86 can not work in x64, you should rebuild it in the x64 environment.</br>
the dll in the branch-master now was build in x86.
# note
the dll now is just avaliable for autoacad2008,since in the file xl2cad_install.reg we registered cad2008 using "[HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\AutoCAD\R17.1\ACAD-6001:804\Applications\xl2cad_cad.Command]"
if you have installded other version of autocad, you should change the "SOFTWARE\Autodesk\AutoCAD\R17.1\ACAD-6001:804\" to fit your cad version.
