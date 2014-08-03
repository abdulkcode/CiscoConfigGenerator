CiscoConfigGenerator
====================

This script will allow the bulk generation of Cisco IOS based configuration files. An excel spreadsheet is used as a front end.

Currently the tool supports the creation of:

* VLANs
* VRFs
* Port-channel interfaces
* Layer 2 settings
* Layer 3 settings
* Static routing
* Prefix-list

Instructions:

1. Download build.xlsx and update to include the relevant configuration
2. Run the script, i.e. ccg.py build.xlsx
