
&AtClient
Procedure Print(Command)
	PrintAtServer();

EndProcedure

&AtServer
Procedure PrintAtServer()
	
	Array = New Array;
	Array.Add(Ref);
	
	Obj = FormAttributeToValue("Object");
	//Template = Obj.GetTemplate("MXLNIR");
	Table = Obj.CreatePrintForm(Array, , "MXLNIR");
	
EndProcedure
