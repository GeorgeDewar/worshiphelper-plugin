﻿using System.Reflection;
using System.Runtime.InteropServices;
using log4net.Config;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("WorshipHelper")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("WorshipHelper")]
[assembly: AssemblyCopyright("Copyright © George Dewar 2020")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("9ef17eed-e3a2-4702-a348-aaf222660dff")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.4.0.*")]
//[assembly: AssemblyFileVersion("1.0.0.0")]

// Reference the configuration file explicitly - otherwise I think the PowerPoint config file gets read
[assembly: XmlConfigurator(ConfigFile = "WorshipHelperVSTO.dll.config", Watch = true)]
