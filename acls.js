var user = [];
user[1]="maritimecb\\os";
user[2]="maritimecb\\sas";
user[3]="maritimecb\\ab";
user[4]="maritimecb\\eo";
user[5]="maritimecb\\sd";
user[6]="maritimecb\\ac";
user[7]="maritimecb\\ik";
user[8]="maritimecb\\js";
user[9]="maritimecb\\ds";
user[10]="maritimecb\\ns";
user[11]="maritimecb\\mg";
user[12]="maritimecb\\vp";

users=12;

var test_i=1;

var wsh = WScript.CreateObject("WScript.Shell");
fso = new ActiveXObject("Scripting.FileSystemObject");
File = fso.GetFile("templ1.txt");
TextStream = File.OpenAsTextStream(1);


str=TextStream.ReadLine();


//-------------------------

//str1=str.split(";");
//WScript.Echo(str1[0].indexOf("\\\\"));
//WScript.Echo(str1[0]+"|"+str1[1]+"|"+str1[2]+"|"+str1[3]);

//---------------------


while (str!="END")

//while (test_i<6)
{
//WScript.Echo(test_i);
	
	str_arr=str.split(";");
	if (str_arr[0].indexOf("\\\\")!=-1){;

		base_path=str_arr[0];	
		if ((str_arr[1]=="+")||(str_arr[1]=="-")||(str_arr[1]=="±")) apath=base_path
		else apath="empty line";
	
	}
	else
	{
		//apath="subdir"
		if(str.indexOf(";;;;***")!=-1) apath="empty line"
		else
		apath=base_path+"\\"+str_arr[0];

	}
	

	
	if (apath!="empty line"){
		
		for(var i = 1; i < users+1; i++){
			
			if (str_arr[i]=="+") {
			command="subinacl.exe /subdirectories \""+apath+"\" /grant="+user[i]+"=f";
			//WScript.Echo(command);
						
			wsh.Run(command,1,true);

			//WScript.Echo("Done");
			
			
			}
			
			if (str_arr[i]=="±") wsh.Run("subinacl.exe /subdirectories \""+apath+"\" /grant="+user[i]+"=r",1,true);
			if (str_arr[i]=="-") wsh.Run("subinacl.exe /subdirectories \""+apath+"\" /revoke="+user[i],1,true);
		;
		};

	}

	
//wsh.Exec("cacls "+source+"\\profiles\\%username%"+" /e /g administrator:f");

	

	//WScript.Echo(apath);
	str=TextStream.ReadLine();
	test_i++;

};
