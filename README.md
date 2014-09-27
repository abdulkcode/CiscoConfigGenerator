=======================
Cisco Config Generator
=======================
This tool will allow the bulk creation of Cisco IOS based configuration using an excel spreadsheet as the front end.  

The intention is to provide engineers with a simple, lightweight utility that allows them to quickly generate configuration without having to manual prepare it in an offline text file.    Worksheets are created which allows the user to enter: 

* Configuration templates
* VLANs
* VRFs
* Port-channel interfaces
* Layer 2 settings
* Layer 3 settings
* Static routing
* Prefix-list

Simply populate whichever worksheet is applicable and then run the script and you're good to go.

Instructions:

1. Download build.xlsx and update to include the relevant configuration
2. Run the script, i.e. ccg.py build.xlsx

==========================
Version Information
==========================
Version 1.x - original release, does not support configuration templates
Version 2.0 - complete re-write of the code, now supports configuration templates with dynamic variables
Version 2.1 - changes to the build worksheet and minor bug fixes

