com2reg  
=======  

This tool creates registry scripts (.reg files) for registering .NET assemblies and native libraries for COM interop.  

Unlike regasm /regfile, com2reg handles registry changes that can made by user-defined methods marked with the COMRegisterFunction attribute.  
  
    Syntax: com2reg AssemblyName [AssemblyName [ ...]] [Options]  
      Options:  
        /regfile:FileName        Generate a reg file with the specified name  
        /codebase                Set the code base in the registry  
        /asmpath:Directory       Look for assembly references here  
        /silent                  Silent mode. Prevents displaying of success messages  
        /? or /help              Display this usage message  
