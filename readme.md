GenericHid
==========

This library is mainly based on Jan Axelsons GenericHid v6.2 (http://janaxelson.com/hidpage.htm#MyExampleCode) with the following modifications:
- the asynchronouos communications part is taken from GenericHid v5.0 to be compatible with .NET 4.0
- some lines / parts of code / ideas are taken from Mike O'Briens HidLibrary (https://github.com/mikeobrien/HidLibrary) and jj-jabbs HidLib (https://github.com/jj-jabb/HidLib)
- the project builds as dedicated generic dll that can also be used in other projects and with other USB devices without any modifications
- devices can be selected not only by their vendor and product ID, but also by the manufacturer or product name and/ or the serial number

For a description of the included files see the comment at the top of HidDevice.cs, except:
- GenericHid.cs doesn't exist any more
- FrmMain.cs renamed to HidDevice.cs and deleted all the form related code
- Debugging.cs and DebuggingDeclarations.cs are not used any more, but still exist
- HidReport.cs offers a class for HID reports

Requirements
------------
.NET 4.0 framework
