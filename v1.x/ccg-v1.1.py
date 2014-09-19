import sys
import os
import datetime

import xlrd             # use pip install xlrd from the c:\python34\scripts directory
from netaddr import *   # use pip install netaddr from the c:\python34\scripts directory

# If pip.exe is not available in the c:\python34\scripts directory, download from:
# http://pip.readthedocs.org/en/latest/installing.html

__author__ = 'Abdul Karim El-Assaad'            # All rights to code reserved 2014
__version__ = '(CCG) Version: 1.1 (18/8/2014)'  # Cisco Config Generator version


#-----------------------------
# Global database variables
#-----------------------------
database = {}       # Main database used to store device, interface, vlans and profiles
variables = {}      # Variables database used to store variables and variable values
errors = {}         # Errors database used to store all errors generated
devices = []        # Only used right at the end to determine which config files were generated
positions = {}      # Store column heading and positions of spreadsheet file

global filename     # The filename used to open the build spreadsheet



#--------------------------------------------------------
# Provide the ability to redirect output to a filename
#--------------------------------------------------------
class Logger(object):
    def __init__(self, filename="Default.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "w")

    def write(self, message):
        #self.terminal.write(message)   # Shows output to screen
        self.log.write(message)         # Writes output to file

#-----------------------------------------------------------------------------------
# This function will read the worksheet and determine the position of each column
#-----------------------------------------------------------------------------------
def locatePos():
    book = xlrd.open_workbook(filename)
    global positions

    #--------------------------------------------------
    # Cycle through each tab in the build spreadsheet
    #--------------------------------------------------
    for sheet in book.sheet_names():
        #-------------------------------------------
        # If it's the instructions sheet ignore it
        #-------------------------------------------
        if sheet == "Instructions":
            continue

        worksheet = book.sheet_by_name(sheet)
        positions[sheet] = {}
        num_columns = worksheet.ncols
        #-------------------------------------------------------------------------------------------
        # Cycle through each column and store the name and it's position in the positions database
        #-------------------------------------------------------------------------------------------
        for x in range(num_columns):
            column_name = worksheet.col(x)[0].value
            positions[sheet][column_name] = x
            #print ("Worksheet: {0} and column name is {1}".format(sheet,column_name))

#----------------------------------------------------------
# Retrieve the column position based on the column name
#----------------------------------------------------------
def GetPos(worksheet,column_name):
    if column_name in positions[worksheet]:
        return positions[worksheet][column_name]
    else:
        print ("ERROR: cannot read %s [%s]"%(worksheet,column_name))
        return -1

#-------------------------------------------------
# Capture log and store it in the error database
#-------------------------------------------------
def CreateDeviceErrorEntry(device_name):

    errors[device_name] = {}
    errors[device_name][0] = {}
    errors[device_name][0]["Location"] = ""
    errors[device_name][0]["ErrorMsg"] = ""
    errors[device_name][0]["Variable"] = ""
    errors[device_name][0]["Action"] = ""
    errors[device_name][0]["Valid"] = False

def CaptureError(worksheet_name,device_name,error_msg, error_variable,action):
    #-------------------------------------------------------------------
    # If there is no entry for this device, create a new record for it
    #-------------------------------------------------------------------
    if device_name not in errors:
        CreateDeviceErrorEntry(device_name)
    #------------------------------------------------------------------
    # Check to see how many error entries are available for this device
    #------------------------------------------------------------------
    for no in errors[device_name]:
        available_number = no
    #-----------------------------------------------------------------
    # If it's the first entry for this device then update the record
    #-----------------------------------------------------------------
    if errors[device_name][0]["Valid"] == False:
        errors[device_name][0]["Location"] = worksheet_name
        errors[device_name][0]["ErrorMsg"] = error_msg
        errors[device_name][0]["Variable"] = error_variable
        errors[device_name][0]["Action"] = action
        errors[device_name][0]["Valid"] = True
    #-------------------------------------------------------------------------------
    # Otherwise create a new entry using the next available number for this device
    #-------------------------------------------------------------------------------
    else:
        available_number = available_number + 1
        errors[device_name][available_number] = {}
        errors[device_name][available_number]["Location"] = worksheet_name
        errors[device_name][available_number]["ErrorMsg"] = error_msg
        errors[device_name][available_number]["Variable"] = error_variable
        errors[device_name][available_number]["Action"] = action
        errors[device_name][available_number]["Valid"] = True

#------------------------------------------------------------------------
# Class used to capture all the functions used to read the configuration
#------------------------------------------------------------------------
class ReadConfig(object):
    def __init__(self):
        #----------------------------------------------------------------
        # Cycle through the columns and store their names and positions
        #----------------------------------------------------------------
        locatePos()

    #------------------------------------
    # Add a new device to the database
    #------------------------------------
    def CreateNewRecord(self,device_name):
        #---------------------------------------------------------------------------------------------
        # If the device name has a space character at the beginning or end fix it and generate a log
        #---------------------------------------------------------------------------------------------
        if " " in device_name:
            device_name = device_name.strip()
            CaptureError("NewRecord",device_name,"N/A","Empty character in name","Auto removed")
        #-----------------------------------------------------------------------
        # If there is already a record for this device, then don't do anything
        #-----------------------------------------------------------------------
        if device_name in database:
            return
        #---------------------------------------------------------------
        # Inform the user that a device has been added to the database
        #---------------------------------------------------------------
        print ("%s device added" % device_name)
        #--------------------------------------------------
        # Define the skeletal record for a device
        #--------------------------------------------------
        profiles = []

        database[device_name] = {}
        database[device_name]["Interface"] = {}
        database[device_name]["Vlans"] = {}
        database[device_name]["VRF"] = {}
        database[device_name]["Routes"] = {}
        database[device_name]["Prefix-Lists"] = {}
        database[device_name]["Profiles"] = profiles

        return

    #----------------------------
    # Add a new VRF to the device
    #-----------------------------
    def CreateNewVrf(self,device_name, vrf_name):
        #---------------------------------------------------------------------------------------------
        # If the device name has a space character at the beginning or end fix it and generate a log
        #---------------------------------------------------------------------------------------------
        if " " in device_name:
            device_name = device_name.strip()
            CaptureError("NewVrf",device_name,"N/A","White space in name","Removed")
        #-----------------------------------------------
        # Make sure that a record exist for the device
        #-----------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #--------------------------------------------------------------------
        # If a VRF with this name already exist for the record quit here
        #--------------------------------------------------------------------
        if vrf_name in database[device_name]["VRF"]:
            CaptureError("NewVrf",device_name,"VRF already exist",vrf_name,"Not used")
            return
        #------------------------------------------------------
        # Otherwise create a skeleton record for the interface
        #--------------------------------------------------------
        else:
            print ("%s will create VRF: [%s]"%(device_name,vrf_name))
            database[device_name]["VRF"][vrf_name] = {}
            database[device_name]["VRF"][vrf_name]["Profile"] = ""
            database[device_name]["VRF"][vrf_name]["RD"] = ""
            database[device_name]["VRF"][vrf_name]["RTImport"] = ""
            database[device_name]["VRF"][vrf_name]["RTExport"] = ""
            return

    #-----------------------------
    # Add a new route for a device
    #-----------------------------
    def CreateNewRoute(self,device_name, route):
        #---------------------------------------------------------------------------------------------
        # If the device name has a space character at the beginning or end fix it and generate a log
        #---------------------------------------------------------------------------------------------
        if " " in device_name:
            device_name = device_name.strip()
            CaptureError("NewRoute",device_name,"N/A","White space in name","Removed")
        #-----------------------------------------------
        # Make sure that a record exist for the device
        #-----------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #------------------------------------------------------------------------
        # If a route with this name already exist for the record quit here
        #------------------------------------------------------------------------
        if route in database[device_name]["Routes"]:
            CaptureError("Routing",device_name,"Duplicate route detected",route,"Ignored")
            return
        #------------------------------------------------------
        # Otherwise create a skeleton record for the interface
        #--------------------------------------------------------
        else:
            print ("%s will create route: [%s]"%(device_name,route))
            database[device_name]["Routes"][route] = {}
            database[device_name]["Routes"][route]["VRF"] = ""
            database[device_name]["Routes"][route]["NextHop"] = ""
            database[device_name]["Routes"][route]["Route"] = ""
            database[device_name]["Routes"][route]["SubnetMask"] = ""
            database[device_name]["Routes"][route]["Name"] = ""
            return

    #------------------------------------
    # Add a new prefix-list for a device
    #------------------------------------
    def CreateNewPrefixList(self,device_name, prefix_name, sequence_number):
        #---------------------------------------------------------------------------------------------
        # If the device name has a space character at the beginning or end fix it and generate a log
        #---------------------------------------------------------------------------------------------
        if " " in device_name:
            device_name = device_name.strip()
            CaptureError("NewPrefix",device_name,"Whitespace in name","N/A","Removed")
        #-----------------------------------------------
        # Make sure that a record exist for the device
        #-----------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #-----------------------------------------------
        # Create a skeleton record for the interface
        #-----------------------------------------------
        print ("%s will create prefix-list: [%s]"%(device_name,prefix_name))
        database[device_name]["Prefix-Lists"][prefix_name] = {}
        database[device_name]["Prefix-Lists"][prefix_name][sequence_number] = {}
        database[device_name]["Prefix-Lists"][prefix_name][sequence_number]["Action"] = ""
        database[device_name]["Prefix-Lists"][prefix_name][sequence_number]["Entry"] = ""
        database[device_name]["Prefix-Lists"][prefix_name][sequence_number]["Valid"] = True
        return

    #-----------------------------------
    # Add a new prefix-list for a device
    #----------------------------------
    def CreateNewPrefixSeq(self,device_name, prefix_name, sequence_number):
        #---------------------------------------------------------------------------------------------
        # If the device name has a space character at the beginning or end fix it and generate a log
        #---------------------------------------------------------------------------------------------
        if " " in device_name:
            device_name = device_name.strip()
            CaptureError("NewPrefix",device_name,"Whitespace in name","N/A","Removed")
        #-----------------------------------------------
        # Make sure that a record exist for the device
        #-----------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #-----------------------------------------------
        # Create a skeleton record for the interface
        #-----------------------------------------------
        database[device_name]["Prefix-Lists"][prefix_name][sequence_number] = {}
        database[device_name]["Prefix-Lists"][prefix_name][sequence_number]["Action"] = ""
        database[device_name]["Prefix-Lists"][prefix_name][sequence_number]["Entry"] = ""
        database[device_name]["Prefix-Lists"][prefix_name][sequence_number]["Valid"] = True
        return


    #--------------------------------------------
    # Associate a new interface to a device
    #--------------------------------------------
    def CreateNewInterface(self,device_name, interface_name):
        #---------------------------------------------------------------------------------------------
        # If the device name has a space character at the beginning or end fix it and generate a log
        #---------------------------------------------------------------------------------------------
        if " " in device_name:
            device_name = device_name.strip()
            CaptureError("NewInterface",device_name,"Whitespace in name","N/A","Removed")
        #-----------------------------------------------
        # Make sure that a record exist for the device
        #-----------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #------------------------------------------------------------------------
        # If an interface with this name already exist for the record quit here
        #------------------------------------------------------------------------
        if interface_name in database[device_name]["Interface"]:
            return
        #------------------------------------------------------
        # Otherwise create a skeleton record for the interface
        #--------------------------------------------------------
        else:
            print ("%s will create interface: [%s]"%(device_name,interface_name))



            database[device_name]["Interface"][interface_name] = {}
            database[device_name]["Interface"][interface_name]["PortProfile1"] = ""
            database[device_name]["Interface"][interface_name]["PortProfile2"] = ""
            database[device_name]["Interface"][interface_name]["MTU"] = ""
            database[device_name]["Interface"][interface_name]["DataVlan"] = ""
            database[device_name]["Interface"][interface_name]["VoiceVlan"] = ""
            database[device_name]["Interface"][interface_name]["Speed"] = ""
            database[device_name]["Interface"][interface_name]["Duplex"] = ""
            database[device_name]["Interface"][interface_name]["Description"] = ""
            database[device_name]["Interface"][interface_name]["NativeVlan"] = ""
            database[device_name]["Interface"][interface_name]["TrunkAllowedVlans"] = ""
            database[device_name]["Interface"][interface_name]["IpAddress"] = ""
            database[device_name]["Interface"][interface_name]["SubnetMask"] = ""
            database[device_name]["Interface"][interface_name]["VRF"] = ""
            database[device_name]["Interface"][interface_name]["PortMakeup"] = ""
            database[device_name]["Interface"][interface_name]["PortType"] = ""
            database[device_name]["Interface"][interface_name]["PortChannelParent"] = ""
            database[device_name]["Interface"][interface_name]["PortChannelGroup"] = ""
            database[device_name]["Interface"][interface_name]["PortChannelMode"] = ""
            database[device_name]["Interface"][interface_name]["PortChannelType"] = ""
            database[device_name]["Interface"][interface_name]["PortChannelProfile"] = ""
            database[device_name]["Interface"][interface_name]["PortEnabled"] = "shutdown"
            #-------------------------------------------------------------------------------------------
            # Determine if it is a physical or logical port as this will help with creation sequence
            #-------------------------------------------------------------------------------------------
            if self.CheckInterfaceMakeup(interface_name,"Physical"):
                database[device_name]["Interface"][interface_name]["PortMakeup"] = "Physical"
            elif self.CheckInterfaceMakeup(interface_name,"Logical"):
                database[device_name]["Interface"][interface_name]["PortMakeup"] = "Logical"

            return

    #------------------------------------
    # Update the interface configuration
    #------------------------------------
    def UpdateInterface(self,device_name,interface_name,update_type,new_config):
        #-----------------------------------------------------------------
        # Strip all leading and trailing white space from the device name
        #-----------------------------------------------------------------
        device_name = device_name.strip()
        #----------------------------------------------------------------
        # If the device doesn't exist, create a new device record for it
        #----------------------------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #----------------------------------------------------------------------
        # If the interface doesn't exist, create a new interface record for it
        #----------------------------------------------------------------------
        if interface_name not in database[device_name]["Interface"]:
            self.CreateNewInterface(device_name,interface_name)
        #--------------------------------------------------------------------------
        # Update the relevant interface parameter with the new configuration value
        #--------------------------------------------------------------------------
        database[device_name]["Interface"][interface_name][update_type] = new_config

    #--------------------------
    # Update the VRF record
    #--------------------------
    def UpdateVrf(self,device_name,vrf_name,update_type,new_config):
        #-----------------------------------------------------------------
        # Strip all leading and trailing white space from the device name
        #-----------------------------------------------------------------
        device_name = device_name.strip()
        #----------------------------------------------------------------
        # If the device doesn't exist, create a new device record for it
        #----------------------------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #----------------------------------------------------------------------
        # If the VRF doesn't exist, create a new VRF record for it
        #----------------------------------------------------------------------
        if vrf_name not in database[device_name]["VRF"]:
            self.CreateNewVrf(device_name,vrf_name)
        #--------------------------------------------------------------------------
        # Update the relevant interface parameter with the new configuration value
        #--------------------------------------------------------------------------
        database[device_name]["VRF"][vrf_name][update_type] = new_config

    #--------------------------
    # Update the route record
    #--------------------------
    def UpdateRoute(self,device_name,route,update_type,new_config):
        #-----------------------------------------------------------------
        # Strip all leading and trailing white space from the device name
        #-----------------------------------------------------------------
        device_name = device_name.strip()
        #----------------------------------------------------------------
        # If the device doesn't exist, create a new device record for it
        #----------------------------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #----------------------------------------------------------------------
        # If the VRF doesn't exist, create a new VRF record for it
        #----------------------------------------------------------------------
        if route not in database[device_name]["Routes"]:
            self.CreateNewRoute(device_name,route)
        #--------------------------------------------------------------------------
        # Update the relevant interface parameter with the new configuration value
        #--------------------------------------------------------------------------
        database[device_name]["Routes"][route][update_type] = new_config


    #--------------------------
    # Update the route record
    #--------------------------
    def UpdatePrefixList(self,device_name,prefix_name,sequence_number,update_type,new_config):
        #-----------------------------------------------------------------
        # Strip all leading and trailing white space from the device name
        #-----------------------------------------------------------------
        device_name = device_name.strip()
        #----------------------------------------------------------------
        # If the device doesn't exist, create a new device record for it
        #----------------------------------------------------------------
        if device_name not in database:
            self.CreateNewRecord(device_name)
        #----------------------------------------------------------------------
        # If the VRF doesn't exist, create a new VRF record for it
        #----------------------------------------------------------------------
        if prefix_name not in database[device_name]["Prefix-Lists"]:
            self.CreateNewPrefixList(device_name,prefix_name,sequence_number)

        #----------------------------------------------------------
        # If the sequence number already exist for this prefix list
        #----------------------------------------------------------
        if sequence_number in database[device_name]["Prefix-Lists"][prefix_name]:
            #-----------------------------------------------
            # And the update_type hasn't already been set
            #-----------------------------------------------
            if database[device_name]["Prefix-Lists"][prefix_name][sequence_number][update_type] == "":
                #---------------------------------------------------------------
                # Update the prefix-list entry with the relevant configuration
                #---------------------------------------------------------------
                database[device_name]["Prefix-Lists"][prefix_name][sequence_number][update_type] = new_config
            #--------------------------------------------------------------------------
            # If the update_type has already been set, then warn about duplicate entry
            #--------------------------------------------------------------------------
            else:
                print ("Duplicate entry already exist")
        return

    #--------------------------------------------------------------
    # Read variables worksheet tab and populate the relevant values
    #--------------------------------------------------------------
    def ReadVariables(self):

        workbook_name = filename
        worksheet_name = "variables"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("-----------------------------------------------------------------")
        print ("Reading '%s' from %s (empty variables ignored)" % (worksheet_name,workbook_name))
        print ("-----------------------------------------------------------------")
        curr_row = -1
        num_rows = worksheet.nrows - 1
        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_variable = GetPos(worksheet_name,"Variable")
        pos_variable_value = GetPos(worksheet_name,"Variable Value")
        #-----------------------------------------------------------------
        # Cycle through the input sheet and load the variable information
        #-----------------------------------------------------------------
        while (curr_row < num_rows):
            curr_row += 1
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------
            variable = worksheet.cell_value(curr_row, pos_variable)
            variable_value = worksheet.cell_value(curr_row, pos_variable_value)
            #-------------------------------------------------------------------------------------------
            # If the cell is empty, a table heading, or has the "!" character in it, then skip the row
            #-------------------------------------------------------------------------------------------
            if curr_row == 0:
                continue
            if not variable:
                continue
            if "!" in variable:
                continue
            if not variable_value:
                continue
            #------------------------------------------------------------------------
            # Read the variable and variable values and update the variable database
            #------------------------------------------------------------------------
            print ("%s variable added to database" % variable)
            variables[variable] = variable_value

    #-----------------------------------------------------
    # Read the vlan worksheet and capture all the values
    #-----------------------------------------------------
    def ReadVlans(self):

        workbook_name = filename
        worksheet_name = "vlans"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("------------------------------------")
        print ("Reading '%s' from %s " % (worksheet_name,workbook_name))
        print ("-----------------------------------")
        current_device = ""
        curr_row = -1
        num_rows = worksheet.nrows - 1
        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_device_name = GetPos(worksheet_name,"Device Name")
        pos_vlan_no = GetPos(worksheet_name,"VLAN No")
        pos_vlan_name = GetPos(worksheet_name,"VLAN Name")
        #----------------------------
        # Loop through all the rows
        #----------------------------
        while curr_row < num_rows:
            curr_row += 1
            #------------------------------------------------------------------------------
            # Bypass the first row (it's a header), has no device name, or no vlan number
            #-------------------------------------------------------------------------------
            if curr_row == 0:
                continue
            elif not worksheet.cell_value(curr_row,pos_device_name):
                continue
            elif not worksheet.cell_value(curr_row,pos_vlan_no):
                continue
            elif "!" in worksheet.cell_value(curr_row,pos_device_name):
                continue
            #---------------------------------------------------------------------------------------
            # Determine if a new device is being processed or whether it's the same as the old one
            #---------------------------------------------------------------------------------------
            if current_device != worksheet.cell_value(curr_row,pos_device_name):
                current_device = worksheet.cell_value(curr_row,pos_device_name)
                self.CreateNewRecord(current_device)
                current_device = current_device.strip()
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------
            vlan = worksheet.cell_value(curr_row,pos_vlan_no)
            vlan_name = worksheet.cell_value(curr_row,pos_vlan_name)
            #--------------------------------------------------------------------------
            # If the VLAN already exist for this device then generate an error message
            #---------------------------------------------------------------------------
            if int(vlan) in database[current_device]["Vlans"]:
                CaptureError("Vlans",current_device,"Vlan already exist",int(vlan),"Not used")
                continue
            #--------------------------------------------------------------------------
            # If no VLAN name is defined, then auto create one and generate a message
            #--------------------------------------------------------------------------
            if not vlan_name:
                vlan_name = "auto-created"
                CaptureError("Vlans",current_device,"Vlan name is not defined",int(vlan), "Auto-created")
            #----------------------------------------------
            # Associate the VLAN with the actual device
            #----------------------------------------------
            if vlan and vlan_name:
                print ("%s will create VLAN [%s]"%(current_device,int(vlan)))
                database[current_device]["Vlans"][vlan] = vlan_name

    #-----------------------------------------------------
    # Read the vrf worksheet and capture all the values
    #-----------------------------------------------------
    def ReadVrf(self):

        workbook_name = filename
        worksheet_name = "vrf"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("---------------------------------")
        print ("Reading '%s' from %s" % (worksheet_name,workbook_name))
        print ("---------------------------------")
        current_device = ""
        curr_row = -1
        num_rows = worksheet.nrows - 1
        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_device_name = GetPos(worksheet_name,"Device Name")
        pos_vrf = GetPos(worksheet_name,"VRF")
        pos_profile = GetPos(worksheet_name,"Profile")
        pos_rd = GetPos(worksheet_name,"RD")
        pos_import_rt = GetPos(worksheet_name,"Import RT  (separated by commas)")
        pos_export_rt = GetPos(worksheet_name,"Export RT  (separated by commas)")
        #----------------------------
        # Loop through all the rows
        #----------------------------
        while curr_row < num_rows:
            curr_row += 1
            #------------------------------------------------------------------------------
            # Bypass the first row (it's a header), has no device name, or no vrf name
            #-------------------------------------------------------------------------------
            if curr_row == 0:
                continue
            elif not worksheet.cell_value(curr_row,pos_device_name):
                continue
            elif not worksheet.cell_value(curr_row,pos_vrf):
                continue
            elif "!" in worksheet.cell_value(curr_row,pos_device_name):
                continue
            #---------------------------------------------------------------------------------------
            # Determine if a new device is being processed or whether it's the same as the old one
            #---------------------------------------------------------------------------------------
            if current_device != worksheet.cell_value(curr_row,pos_device_name):
                current_device = worksheet.cell_value(curr_row,pos_device_name)
                self.CreateNewRecord(current_device)
                current_device = current_device.strip()
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------
            vrf = worksheet.cell_value(curr_row,pos_vrf)
            profile = worksheet.cell_value(curr_row,pos_profile)
            rd = worksheet.cell_value(curr_row,pos_rd)
            import_rt = worksheet.cell_value(curr_row,pos_import_rt)
            export_rt = worksheet.cell_value(curr_row,pos_export_rt)
            #--------------------------------------------------------------------------
            # If the VRF already exist for this device then generate an error message
            #---------------------------------------------------------------------------
            if vrf in database[current_device]["VRF"]:
                CaptureError("Vrf",current_device,"VRF already exist",vrf,"Not used")
                continue
            #------------------------------------------------------
            # Update the VRF record with the relevant information
            #------------------------------------------------------
            if profile:
                self.UpdateVrf(current_device,vrf,"Profile", profile)
            if rd:
                self.UpdateVrf(current_device,vrf,"RD", rd)
            if import_rt:
                self.UpdateVrf(current_device,vrf,"RTImport", import_rt)
            if export_rt:
                self.UpdateVrf(current_device,vrf,"RTExport", export_rt)

    #-----------------------------------------------------
    # Read the routing worksheet and capture all the values
    #-----------------------------------------------------
    def ReadRouting(self):

        workbook_name = filename
        worksheet_name = "routing"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("---------------------------------")
        print ("Reading '%s' from %s" % (worksheet_name,workbook_name))
        print ("---------------------------------")
        current_device = ""
        curr_row = -1
        num_rows = worksheet.nrows - 1
        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_device_name = GetPos(worksheet_name,"Device Name")
        pos_vrf = GetPos(worksheet_name,"VRF (leave blank if global)")
        pos_route = GetPos(worksheet_name,"Route (x.x.x.x/x)")
        pos_next_hop = GetPos(worksheet_name,"Next Hop")
        pos_route_name = GetPos(worksheet_name,"Route Name (no spaces)")
        #----------------------------
        # Loop through all the rows
        #----------------------------
        while curr_row < num_rows:
            curr_row += 1
            #------------------------------------------------------------------------------
            # Bypass the first row (it's a header), has no device name, or no route entry
            #-------------------------------------------------------------------------------
            if curr_row == 0:
                continue
            elif not worksheet.cell_value(curr_row,pos_device_name):
                continue
            elif not worksheet.cell_value(curr_row,pos_route):
                continue
            elif "!" in worksheet.cell_value(curr_row,pos_device_name):
                continue
            #---------------------------------------------------------------------------------------
            # Determine if a new device is being processed or whether it's the same as the old one
            #---------------------------------------------------------------------------------------
            if current_device != worksheet.cell_value(curr_row,pos_device_name):
                current_device = worksheet.cell_value(curr_row,pos_device_name)
                self.CreateNewRecord(current_device)
                current_device = current_device.strip()
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------
            vrf = worksheet.cell_value(curr_row,pos_vrf)
            route = worksheet.cell_value(curr_row,pos_route)
            route_next_hop = worksheet.cell_value(curr_row,pos_next_hop)
            route_name = worksheet.cell_value(curr_row,pos_route_name)
            #--------------------------------------------------------------------------
            # If the route already exist for this device then generate an error message
            #---------------------------------------------------------------------------
            if route in database[current_device]["Routes"]:
                CaptureError("Routing",current_device,"Route already exists", route,"Skipped")
                continue
            #------------------------------------------------------
            # Update the route record with the relevant information
            #------------------------------------------------------
            if route:
                self.UpdateRoute(current_device,route,"Routes", route)
            if vrf:
                self.UpdateRoute(current_device,route,"VRF", vrf)
            if route_next_hop:
                self.UpdateRoute(current_device,route,"NextHop", route_next_hop)
            if route_name:
                self.UpdateRoute(current_device,route,"Name", route_name)

    #-----------------------------------------------------
    # Read the routing worksheet and capture all the values
    #-----------------------------------------------------
    def ReadPrefixList(self):

        workbook_name = filename
        worksheet_name = "prefix-list"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("---------------------------------")
        print ("Reading '%s' from %s" % (worksheet_name,workbook_name))
        print ("---------------------------------")
        current_device = ""
        curr_row = -1
        num_rows = worksheet.nrows - 1
        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_device_name = GetPos(worksheet_name,"Device Name")
        pos_prefix_name = GetPos(worksheet_name,"Prefix-List Name")
        pos_prefix_seq = GetPos(worksheet_name,"Prefix-List Sequence No")
        pos_prefix_action = GetPos(worksheet_name,"Prefix-List Action (permit/deny)")
        pos_prefix_entry = GetPos(worksheet_name,"Prefix-List Entry")
        #----------------------------
        # Loop through all the rows
        #----------------------------
        while curr_row < num_rows:
            curr_row += 1
            #------------------------------------------------------------------------------
            # Bypass the first row (it's a header), has no device name, or no route entry
            #-------------------------------------------------------------------------------
            if curr_row == 0:
                continue
            elif not worksheet.cell_value(curr_row,pos_device_name):
                continue
            elif not worksheet.cell_value(curr_row,pos_prefix_name):
                continue
            elif not worksheet.cell_value(curr_row,pos_prefix_seq):
                continue
            elif "!" in worksheet.cell_value(curr_row,pos_device_name):
                continue
            #---------------------------------------------------------------------------------------
            # Determine if a new device is being processed or whether it's the same as the old one
            #---------------------------------------------------------------------------------------
            if current_device != worksheet.cell_value(curr_row,pos_device_name):
                current_device = worksheet.cell_value(curr_row,pos_device_name)
                self.CreateNewRecord(current_device)
                current_device = current_device.strip()
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------
            prefix_name = worksheet.cell_value(curr_row,pos_prefix_name)
            prefix_seq = int(worksheet.cell_value(curr_row,pos_prefix_seq))
            prefix_action = worksheet.cell_value(curr_row,pos_prefix_action)
            prefix_entry = worksheet.cell_value(curr_row,pos_prefix_entry)
            #--------------------------------------------------------------------
            # Create a new prefix for the device if it does not already exists
            #--------------------------------------------------------------------
            if prefix_name not in database[current_device]["Prefix-Lists"]:
                self.CreateNewPrefixList(current_device,prefix_name,prefix_seq)
            #----------------------------------------------------------------------------------------
            # Create a new sequence number for an existing prefix-list if it does not already exist
            #-------------------------------------------------------------------------------------------
            if prefix_seq not in database[current_device]["Prefix-Lists"][prefix_name]:
                self.CreateNewPrefixSeq(current_device,prefix_name,prefix_seq)
            #------------------------------------------------------------------------------
            # Do some basic checks before updating the action for the prefix-list sequence
            #-------------------------------------------------------------------------------
            if prefix_action:
                #-----------------------------------------------------------------
                # If the action for this sequence number is empty, then update it
                #-----------------------------------------------------------------
                if database[current_device]["Prefix-Lists"][prefix_name][prefix_seq]["Action"] == "":
                    self.UpdatePrefixList(current_device,prefix_name,prefix_seq,"Action", prefix_action)
                #------------------------------------------------------------------------
                # Otherwise ignore it and generate an error, duplicate has been detected
                #------------------------------------------------------------------------
                else:
                    CaptureError("PrefixList",current_device,"Duplicate prefix action",prefix_name,"Not used")
                    continue
            #------------------------------------------------------------------------------
            # Do some basic checks before updating the entry for the prefix-list sequence
            #-------------------------------------------------------------------------------
            if prefix_entry:
                #-----------------------------------------------------------------
                # If the entry for this sequence number is empty, then update it
                #-----------------------------------------------------------------
                if database[current_device]["Prefix-Lists"][prefix_name][prefix_seq]["Entry"] == "":
                    self.UpdatePrefixList(current_device,prefix_name,prefix_seq,"Entry", prefix_entry)
                #------------------------------------------------------------------------
                # Otherwise ignore it and generate an error, duplicate has been detected
                #------------------------------------------------------------------------
                else:
                    CaptureError("PrefixList",current_device,"Duplicate prefix entry",prefix_name,"Not used")
                    continue


    #-------------------------------------------------------------------------
    # Read the profiles worksheet and store the data in the database
    #--------------------------------------------------------------------------
    def ReadProfiles(self):

        workbook_name = filename
        worksheet_name = "profiles"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("---------------------------------")
        print ("Reading '%s' from %s" % (worksheet_name,workbook_name))
        print ("---------------------------------")
        current_device = ""
        curr_row = -1
        num_rows = worksheet.nrows - 1
        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_device_name = GetPos(worksheet_name,"Device Name")
        pos_profile = GetPos(worksheet_name,"Profile")
        #----------------------------
        # Loop through all the rows
        #----------------------------
        while curr_row < num_rows:
            curr_row += 1
            #--------------------------------------------------------
            # Bypass the first row (it's a header) and emtpy rows
            #--------------------------------------------------------
            if curr_row == 0:
                continue
            elif not worksheet.cell_value(curr_row,pos_device_name):
                continue
            elif not worksheet.cell_value(curr_row,pos_profile):
                continue
            elif "!" in worksheet.cell_value(curr_row,pos_device_name):
                continue
            #----------------------------------------------------------------------------------------
            # Determine if a new device is being processed or whether it's the same as the old one
            #----------------------------------------------------------------------------------------
            if current_device != worksheet.cell_value(curr_row,pos_device_name):
                current_device = worksheet.cell_value(curr_row,pos_device_name)
                self.CreateNewRecord(current_device)
                current_device = current_device.strip()
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------
            profile = worksheet.cell_value(curr_row,pos_profile)
            #--------------------------------------------------------------
            # Check to see if the profile is referencing a valid variable
            #--------------------------------------------------------------
            if profile in variables:
                    #----------------------------------------------------------------------------
                    # If the profile hasn't already been configured for this device, then add it
                    #----------------------------------------------------------------------------
                    if profile not in database[current_device]["Profiles"]:
                        print ("%s will use profile: [%s]"%(current_device,profile))
                        database[current_device]["Profiles"].append(profile)
                    #--------------------------------------------------------------
                    # Otherwise it must already be configured, report a duplicate
                    #--------------------------------------------------------------
                    else:
                        CaptureError("Profiles",current_device,"Device already using profile",profile,"Not used")
            #------------------------------------------------------------------------
            # Referencing a variable that does not exist, ignore and generate error
            #------------------------------------------------------------------------
            else:
                CaptureError("Profiles",current_device,"Invalid/blank profile used", profile,"Not used")


    #----------------------------------------------------------
    # Read the portchannel worksheet and capture all the values
    #----------------------------------------------------------
    def ReadPortChannel(self):

        workbook_name = filename
        worksheet_name = "portchannels"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("-------------------------------------")
        print ("Reading '%s' from %s" % (worksheet_name,workbook_name))
        print ("-------------------------------------")

        current_device = ""
        curr_row = -1
        num_rows = worksheet.nrows - 1

        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_device_name = GetPos(worksheet_name,"Device Name")
        pos_pc_interface = GetPos(worksheet_name,"Interface")
        pos_pc_group = GetPos(worksheet_name,"Port-Channel Group")
        pos_pc_mode = GetPos(worksheet_name,"Port-Channel Mode (active/on/etc)")
        pos_pc_type = GetPos(worksheet_name,"Port-Channel Type (layer2 or layer3)")
        pos_pc_profile = GetPos(worksheet_name,"Port-Channel Profile")
        pos_pc_members = GetPos(worksheet_name,"Port-Channel Members (separated by commas)")
        pos_pc_description = GetPos(worksheet_name,"Description")
        #----------------------------
        # Loop through all the rows
        #----------------------------
        while curr_row < num_rows:
            curr_row += 1
            #-----------------------------------------------------
            # Bypass the first row (it's a header) and emtpy rows
            #-----------------------------------------------------
            if curr_row == 0:
                continue
            elif not worksheet.cell_value(curr_row,pos_device_name):
                continue
            elif "!" in worksheet.cell_value(curr_row,pos_device_name):
                continue
            #----------------------------------------------------------------------------------------
            # Determine if a new device is being processed or whether it's the same as the old one
            #----------------------------------------------------------------------------------------
            if current_device != worksheet.cell_value(curr_row,pos_device_name):
                current_device = worksheet.cell_value(curr_row,pos_device_name)
                self.CreateNewRecord(current_device)
                current_device = current_device.strip()
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------
            interface = worksheet.cell_value(curr_row,pos_pc_interface)
            port_group = worksheet.cell_value(curr_row, pos_pc_group)
            port_mode = worksheet.cell_value(curr_row, pos_pc_mode)
            port_profile = worksheet.cell_value(curr_row, pos_pc_profile)
            port_type = worksheet.cell_value(curr_row, pos_pc_type).strip().title()
            description = worksheet.cell_value(curr_row,pos_pc_description)
            #--------------------------------------------------------------------------------
            # Capitalise the the first letter of the interface name to avoid database errors
            #--------------------------------------------------------------------------------
            if interface:
                interface = interface.title()
            #----------------------------------------------------------------------------------
            # If no interface entry exists then create a new record and update the description
            #----------------------------------------------------------------------------------
            if interface not in database[current_device]["Interface"]:
                self.CreateNewInterface(current_device,interface)
                self.UpdateInterface(current_device,interface,"PortType", port_type)
            #-----------------------------------------------------------------------------------
            # Update the actual interface record with details from the portchannels worksheet
            #-----------------------------------------------------------------------------------
            if port_group:
                database[current_device]["Interface"][interface]["PortChannelProfile"] = port_group
            if port_mode:
                database[current_device]["Interface"][interface]["PortChannelMode"] = port_mode
            if port_type:
                database[current_device]["Interface"][interface]["PortChannelType"] = port_type
            if port_profile:
                database[current_device]["Interface"][interface]["PortChannelProfile"] = port_profile
            #--------------------------------------------------------------------------------------------------
            # If the port-channel interface has had it's description defined elsewhere, then don't overwrite
            #--------------------------------------------------------------------------------------------------
            if description:
                if database[current_device]["Interface"][interface]["Description"]:
                    CaptureError("PortChannel",current_device,"Description already defined",interface+":"+description,"Ignored")
                #==========================================================================================
                # If the interface has no description defined so far, then use the one from this worksheet
                #==========================================================================================
                else:
                    self.UpdateInterface(current_device,interface,"Description", description)
            #----------------------------------------------------------------------------------------------------------
            #Cycle through the member interface of each port-channel and determine whether there is an entry for them
            #----------------------------------------------------------------------------------------------------------
            member_interfaces = worksheet.cell_value(curr_row,pos_pc_members)
            #------------------------------------------------------------------------------------------------
            # Make sure that the member interfaces actually exist and if so, configure the channel-group info
            #------------------------------------------------------------------------------------------------
            if member_interfaces:
                for member_port in member_interfaces.split(','):
                    print ("Processing member: %s"%member_port)
                    #-------------------------------------------------------------------
                    # If there is a space in the port-channel members, strip the space
                    #-------------------------------------------------------------------
                    if " " in member_port:
                        member_port = member_port.replace(" ","")
                    #-----------------------------------------------------------------------------------------------
                    # Could be an error if someone puts a , without any value after it, prevent this from happening
                    #-----------------------------------------------------------------------------------------------
                    if not member_port:
                        CaptureError("PortChan",current_device,"Invalid member after ',' (blank)",interface,"Member Ignored")
                        continue
                    #------------------------------------------
                    # Make sure the first letter it capitalised
                    #------------------------------------------
                    member_port = member_port.title()
                    #-----------------------------------------------------------------
                    # If the member interface has not been defined yet, create it
                    #------------------------------------------------------------------
                    if member_port not in database[current_device]["Interface"]:
                        channel_group = worksheet.cell_value(curr_row,pos_pc_group)
                        channel_mode = worksheet.cell_value(curr_row,pos_pc_mode)

                        self.CreateNewInterface(current_device,member_port)
                        self.UpdateInterface(current_device,member_port,"PortType", port_type)
                        self.UpdateInterface(current_device,member_port,"PortChannelParent", interface)
                        self.UpdateInterface(current_device,member_port,"PortChannelGroup", int(channel_group))
                        self.UpdateInterface(current_device,member_port,"PortChannelMode", channel_mode)
                        self.UpdateInterface(current_device,member_port,"PortEnabled", "Yes")



    #-----------------------------------------------------
    # Read the layer2 worksheet and capture all the values
    #-----------------------------------------------------
    def ReadLayer2(self):

        workbook_name = filename
        worksheet_name = "layer2"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("-------------------------------------")
        print ("Reading '%s' from %s" % (worksheet_name,workbook_name))
        print ("-------------------------------------")
        current_device = ""
        curr_row = -1
        num_rows = worksheet.nrows - 1

        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_device_name = GetPos(worksheet_name,"Device Name")
        pos_port = GetPos(worksheet_name,"Port")
        pos_port_enabled = GetPos(worksheet_name,"Port Enabled (yes/no)")
        pos_port_profile1 = GetPos(worksheet_name,"Port Profile 1")
        pos_port_profile2 = GetPos(worksheet_name,"Port Profile 2")
        pos_port_mtu = GetPos(worksheet_name,"MTU (leave blank for default)")
        pos_port_data_vlan = GetPos(worksheet_name,"Data VLAN")
        pos_port_voice_vlan = GetPos(worksheet_name,"Voice VLAN")
        pos_port_speed = GetPos(worksheet_name,"Speed")
        pos_port_duplex = GetPos(worksheet_name,"Duplex")
        pos_port_description = GetPos(worksheet_name,"Description")
        pos_port_trunk = GetPos(worksheet_name,"Allowed VLANs (separated by commas)")
        pos_port_native = GetPos(worksheet_name,"Native VLAN")
        #----------------------------
        # Loop through all the rows
        #----------------------------
        while curr_row < num_rows:
            curr_row += 1
            #------------------------------------------------------
            # Bypass the first row (it's a header) and emtpy rows
            #------------------------------------------------------
            if curr_row == 0:
                continue
            elif not worksheet.cell_value(curr_row,pos_device_name):
                continue
            elif not worksheet.cell_value(curr_row,pos_port):
                continue
            elif "!" in worksheet.cell_value(curr_row,pos_device_name):
                continue
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------

            interface = worksheet.cell_value(curr_row,pos_port)
            port_profile1 = worksheet.cell_value(curr_row, pos_port_profile1)
            port_profile2 = worksheet.cell_value(curr_row, pos_port_profile2)
            port_enabled = worksheet.cell_value(curr_row, pos_port_enabled)
            description = str(worksheet.cell_value(curr_row,pos_port_description))
            speed = worksheet.cell_value(curr_row,pos_port_speed)
            duplex = worksheet.cell_value(curr_row,pos_port_duplex)
            mtu = str(worksheet.cell_value(curr_row,pos_port_mtu))
            data_vlan = worksheet.cell_value(curr_row, pos_port_data_vlan)
            voice_vlan = worksheet.cell_value(curr_row, pos_port_voice_vlan)
            trunk_allowed_vlans = worksheet.cell_value(curr_row,pos_port_trunk)
            native_vlan = worksheet.cell_value(curr_row,pos_port_native)
            #---------------------------------------------------------------------------------------
            # Determine if a new device is being processed or whether it's the same as the old one
            #---------------------------------------------------------------------------------------
            if current_device != worksheet.cell_value(curr_row,pos_device_name):
                current_device = worksheet.cell_value(curr_row,pos_device_name)
                self.CreateNewRecord(current_device)
                current_device = current_device.strip()
            #--------------------------------------------------------------------------------
            # Capitalise the the first letter of the interface name to avoid database errors
            #--------------------------------------------------------------------------------
            if interface:
                interface = interface.title()
            #-------------------------------------------------------
            # Create a new record for this interface
            #-------------------------------------------------------
            self.CreateNewInterface(current_device,interface)
            #-------------------------------------------------------
            # Define the interface as a Layer 2 port
            #-------------------------------------------------------
            self.UpdateInterface(current_device,interface,"PortType", "Layer2")
            #-------------------------------------------------------
            # Update the various parameters for this interface
            #-------------------------------------------------------
            if port_profile1:
                self.UpdateInterface(current_device,interface,"PortProfile1", port_profile1)
            if port_profile2:
                self.UpdateInterface(current_device,interface,"PortProfile2", port_profile2)
            if port_enabled:
                self.UpdateInterface(current_device,interface,"PortEnabled", port_enabled)
            if mtu:
                self.UpdateInterface(current_device,interface,"MTU", mtu)
            if data_vlan:
                self.UpdateInterface(current_device,interface,"DataVlan", int(data_vlan))
            if voice_vlan:
                self.UpdateInterface(current_device,interface,"VoiceVlan", int(voice_vlan))
            if speed:
                self.UpdateInterface(current_device,interface,"Speed", speed)
            if duplex:
                self.UpdateInterface(current_device,interface,"Duplex", duplex)
            if trunk_allowed_vlans:
                self.UpdateInterface(current_device,interface,"TrunkAllowedVlans", trunk_allowed_vlans)
            if native_vlan:
                self.UpdateInterface(current_device,interface,"NativeVlan", int(native_vlan))


            if description:
                if database[current_device]["Interface"][interface]["Description"]:
                    CaptureError("Layer2",current_device,"Description already defined",interface+":"+description,"Ignored")
                else:
                    self.UpdateInterface(current_device,interface,"Description", description)


    #--------------------------------------------------------------
    # Read Layer 3 worksheet tab and populate the relevant values
    #--------------------------------------------------------------
    def ReadLayer3(self):

        workbook_name = filename
        worksheet_name = "layer3"
        workbook = xlrd.open_workbook(workbook_name)
        worksheet = workbook.sheet_by_name(worksheet_name)
        print ("-------------------------------------")
        print ("Reading '%s' from %s" % (worksheet_name,workbook_name))
        print ("-------------------------------------")
        current_device = ""
        curr_row = -1
        num_rows = worksheet.nrows - 1

        #-----------------------------------
        # Grab the position of each column
        #-----------------------------------
        pos_device_name = GetPos(worksheet_name,"Device Name")
        pos_interface = GetPos(worksheet_name,"Interface")
        pos_port_enabled = GetPos(worksheet_name,"Port Enabled (yes/no)")
        pos_port_profile1 = GetPos(worksheet_name,"Port Profile 1")
        pos_port_profile2 = GetPos(worksheet_name,"Port Profile 2")
        pos_vrf = GetPos(worksheet_name,"VRF (leave blank if global)")
        pos_ip_address = GetPos(worksheet_name,"IP Address (x.x.x.x/x)")
        pos_mtu = GetPos(worksheet_name,"MTU (leave blank for default)")
        pos_description = GetPos(worksheet_name,"Description")

        #----------------------------
        # Loop through all the rows
        #----------------------------
        while curr_row < num_rows:
            curr_row += 1
            #------------------------------------------------------
            # Bypass the first row (it's a header) and emtpy rows
            #------------------------------------------------------
            if curr_row == 0:
                continue
            elif not worksheet.cell_value(curr_row,pos_device_name):
                continue
            elif not worksheet.cell_value(curr_row,pos_interface):
                continue
            elif "!" in worksheet.cell_value(curr_row,pos_device_name):
                continue
            #-------------------------------------------------------------------------
            # Retrieve the content of each cell and store it in a local variable name
            #-------------------------------------------------------------------------
            interface = worksheet.cell_value(curr_row,pos_interface)
            port_profile1 = worksheet.cell_value(curr_row, pos_port_profile1)
            port_profile2 = worksheet.cell_value(curr_row, pos_port_profile2)
            description = str(worksheet.cell_value(curr_row,pos_description))
            ip_address = worksheet.cell_value(curr_row,pos_ip_address)
            mtu = str(worksheet.cell_value(curr_row,pos_mtu))
            port_enabled = worksheet.cell_value(curr_row, pos_port_enabled)
            vrf = str(worksheet.cell_value(curr_row,pos_vrf))
            #---------------------------------------------------------------------------------------
            # Determine if a new device is being processed or whether it's the same as the old one
            #---------------------------------------------------------------------------------------
            if current_device != worksheet.cell_value(curr_row,pos_device_name):
                current_device = worksheet.cell_value(curr_row,pos_device_name)
                self.CreateNewRecord(current_device)
                current_device = current_device.strip()
            #--------------------------------------------------------------------------------
            # Capitalise the the first letter of the interface name to avoid database errors
            #--------------------------------------------------------------------------------
            if interface:
                interface = interface.title()
            #-------------------------------------------------------
            # Try and create a new record for this interface
            #-------------------------------------------------------
            self.CreateNewInterface(current_device,interface)
            #-------------------------------------------------------
            # Define the interface as a Layer 3 port
            #-------------------------------------------------------
            self.UpdateInterface(current_device,interface,"PortType", "Layer3")
            #-------------------------------------------------------
            # Update the various parameters for this interface
            #-------------------------------------------------------
            if port_enabled:
                self.UpdateInterface(current_device,interface,"PortEnabled", port_enabled)
            if port_profile1:
                self.UpdateInterface(current_device,interface,"PortProfile1", port_profile1)
            if port_profile2:
                self.UpdateInterface(current_device,interface,"PortProfile2", port_profile2)
            if vrf:
                self.UpdateInterface(current_device,interface,"VRF", vrf)
            if ip_address:
                self.UpdateInterface(current_device,interface,"IpAddress", ip_address)
            if mtu:
                self.UpdateInterface(current_device,interface,"MTU", mtu)

            if description:
                if database[current_device]["Interface"][interface]["Description"]:
                    CaptureError("Layer3",current_device,"Description already defined",interface+":"+description,"Ignored")
                else:
                    self.UpdateInterface(current_device,interface,"Description", description)

    #------------------------------------------------------
    # Check to see if the interface is logical or physical
    #------------------------------------------------------
    def CheckInterfaceMakeup(self,interface_name,interface_type):
        #----------------------------------------------------------------------------
        # Discard logical interface types if you're looking for a physical interface
        #----------------------------------------------------------------------------
        if interface_type == "Physical":
            if "Po" in interface_name:
                return False
            elif "Tu" in interface_name:
                return False
            elif "Lo" in interface_name:
                return False
            elif "Vl" in interface_name:
                return False
            else:
                return True
        #----------------------------------------------------------------
        # Show the interface if it has logical interface characteristics
        #----------------------------------------------------------------
        if interface_type == "Logical":
            if "Po" in interface_name:
                return True
            elif "Tu" in interface_name:
                return True
            elif "Lo" in interface_name:
                return True
            elif "Vl" in interface_name:
                return True
            else:
                return False


#------------------------------------------------------------------
# Class used for actually creating all the relevant configuration
#------------------------------------------------------------------
class GenerateConfig(object):
    def __init__(self):
        pass

    #--------------------------------------------------------------------------------------
    # Check the record for VRF errors and correct them prior to generating configuration
    #--------------------------------------------------------------------------------------
    def CheckVrfErrors(self,device_name,vrf_name):
        #---------------------------------------------------------------------------
        # Basic check to make sure the device and interface exist before proceeding
        #---------------------------------------------------------------------------
        if device_name not in database:
            return
        if vrf_name not in database[device_name]["VRF"]:
            return
        #----------------------------------------------
        # Store the values in easier to read variables
        #----------------------------------------------
        vrf_rd = database[device_name]["VRF"][vrf_name]["RD"]
        vrf_profile = database[device_name]["VRF"][vrf_name]["Profile"]
        #---------------------------------------------------------------
        # If there is no RD defined, then don't bother creating the VRF
        #---------------------------------------------------------------
        if not vrf_rd:
            database[device_name]["VRF"][vrf_name]["Profile"] = ""
            database[device_name]["VRF"][vrf_name]["RTImport"] = ""
            database[device_name]["VRF"][vrf_name]["RTExport"] = ""
            CaptureError("Vrf",device_name,"VRF has no RD defined", vrf_name,"Not used")
        #---------------------------------------------------------------------
        # Check to see if the VRF profile actually exist in the variable tab
        #---------------------------------------------------------------------
        if vrf_profile and vrf_profile not in variables:
            database[device_name]["VRF"][vrf_name]["Profile"] = ""
            CaptureError("Vrf",device_name,"Invalid/blank profile for VRF", vrf_name+":"+vrf_profile,"Not used")


    #--------------------------------------------------------------------------------------
    # Check the record for routing errors and correct them prior to generating configuration
    #--------------------------------------------------------------------------------------
    def CheckRoutingErrors(self,device_name,route):
        #---------------------------------------------------------------------------
        # Basic check to make sure the device and interface exist before proceeding
        #---------------------------------------------------------------------------
        if device_name not in database:
            return
        if route not in database[device_name]["Routes"]:
            return
        if not database[device_name]["Routes"][route]["NextHop"]:
            CaptureError("Routing",device_name,"No nexthop defined", route,"Not used")
            return
        #-----------------------------------------------------------------------------------
        # Check to see if the VRF has been created on the switch via the "vrf" Tab
        #-----------------------------------------------------------------------------------
        if database[device_name]["Routes"][route]["VRF"]:
            vrf = database[device_name]["Routes"][route]["VRF"]
            if vrf not in database[device_name]["VRF"]:
                print ("  !-- VRF has not been defined in the 'vrf' tab for this device")
                CaptureError("Routing",device_name,"VRF for route does not exist",route+":"+vrf,"Warning generated")
        #-------------------------------------------
        # Check to see if it's a valid route syntax
        #-------------------------------------------
        if "/" in route:
                try:
                    #----------------------------------------------------------------------------------
                    # It's valid, populate the records with the correct information
                    #----------------------------------------------------------------------------------
                    full_route = IPNetwork(route)
                    database[device_name]["Routes"][route]["Route"] = str(full_route.ip)
                    database[device_name]["Routes"][route]["SubnetMask"] = str(full_route.netmask)
                    #-------------------------------------------------------------------
                    # Check to see if there is a space in the route name if so, trim it
                    #-------------------------------------------------------------------
                    route_name = database[device_name]["Routes"][route]["Name"]
                    if route_name:
                        if " " in route_name:
                            fix_name = route_name.replace(" ","")
                            database[device_name]["Routes"][route]["Name"] = fix_name
                            CaptureError("Routing",device_name,"Space found in route name",route_name,"Space removed")
                #--------------------------------------------------------------------------------------------
                # It's not valid, generate an error message and clear the record so no config is generated
                #--------------------------------------------------------------------------------------------
                except:
                    print ("!-- invalid route, config not generated")
                    CaptureError("Routing",device_name,route,"Invalid syntax","Not used")
                    database[device_name]["Routes"][route]["Route"] = ""
                    database[device_name]["Routes"][route]["VRF"] = ""
                    database[device_name]["Routes"][route]["SubnetMask"] = ""
                    database[device_name]["Routes"][route]["NextHop"] = ""
                    database[device_name]["Routes"][route]["Name"] = ""


    #--------------------------------------------------------------------------------------
    # Check the record for prefix-list errors and correct them prior to generating configuration
    #--------------------------------------------------------------------------------------
    def CheckPrefixErrors(self,device_name,prefix_name):
        #---------------------------------------------------------------------------
        # Basic check to make sure the device and interface exist before proceeding
        #---------------------------------------------------------------------------
        if device_name not in database:
            return
        if prefix_name not in database[device_name]["Prefix-Lists"]:
            return
        #------------------------------------------------------
        # Cycle through each sequence number of the prefix-list
        #------------------------------------------------------
        for sequence in database[device_name]["Prefix-Lists"][prefix_name]:
            prefix_action = database[device_name]["Prefix-Lists"][prefix_name][sequence]["Action"]
            prefix_entry = database[device_name]["Prefix-Lists"][prefix_name][sequence]["Entry"]
            #------------------------------------------------------------------------------
            # If the prefix action or prefix entry is empty, declare the record as invalid
            #------------------------------------------------------------------------------
            if not prefix_action or not prefix_entry:
                database[device_name]["Prefix-Lists"][prefix_name][sequence]["Valid"] = False
                CaptureError("PrefixList",device_name,"Invalid prefix-list seq",("Seq:"+str(sequence)),"Not used")

    #-------------------------------------------------------------------------------------------
    # Check the record for interface errors and correct them prior to generating configuration
    #------------------------------------------------------------------------------------------
    def CheckInterfaceErrors(self,device_name,interface_name):
        #---------------------------------------------------------------------------
        # Basic check to make sure the device and interface exist before proceeding
        #---------------------------------------------------------------------------
        if device_name not in database:
            return
        if interface_name not in database[device_name]["Interface"]:
            return
        #---------------------------------------------------------------------------------------
        # If default values are used, no point in generating configuration, zero out the values
        #---------------------------------------------------------------------------------------
        if ("auto" or "Auto") in database[device_name]["Interface"][interface_name]["Speed"]:
            database[device_name]["Interface"][interface_name]["Speed"] = ""
        if ("auto" or "Auto") in database[device_name]["Interface"][interface_name]["Duplex"]:
            database[device_name]["Interface"][interface_name]["Duplex"] = ""
        if "1500" in database[device_name]["Interface"][interface_name]["MTU"]:
            database[device_name]["Interface"][interface_name]["MTU"] = ""
        #---------------------------------------------------------------------------
        # If the port is configured as a trunk and data access port, use only TRUNK
        #----------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["TrunkAllowedVlans"] and database[device_name]["Interface"][interface_name]["DataVlan"]:
            print ("  !- Trunk and access port inconsistency, converting to trunk only")
            CaptureError("Layer2",device_name,"Trunk+Data VLAN configured",interface_name,"Only trunk used")
            database[device_name]["Interface"][interface_name]["DataVlan"] = ""
            database[device_name]["Interface"][interface_name]["VoiceVlan"] = ""
        #---------------------------------------------------------------------------
        # If the port is configured as a trunk and voice access port, use only TRUNK
        #---------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["TrunkAllowedVlans"] and database[device_name]["Interface"][interface_name]["VoiceVlan"]:
            print ("  !- Trunk and access port inconsistency, converting to trunk only")
            CaptureError("Layer2",device_name,"Trunk+Voice VLAN configured",interface_name,"Only trunk used")
            database[device_name]["Interface"][interface_name]["DataVlan"] = ""
            database[device_name]["Interface"][interface_name]["VoiceVlan"] = ""
        #------------------------------------------------------------------------------
        # Check the IP address syntax and convert x.x.x.x/y to x.x.x.x y.y.y.y format
        #------------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["IpAddress"]:
            ip_address = database[device_name]["Interface"][interface_name]["IpAddress"]
            if ip_address and "/" in ip_address:
                try:
                    #----------------------------------------------------------------------------------
                    # A valid IP address syntax and host/subnet detected, allow configuration
                    #----------------------------------------------------------------------------------
                    addr = IPNetwork(ip_address)
                    ip,mask = ip_address.split("/")

                    if IPAddress(ip) in IPNetwork(ip_address).iter_hosts():
                        database[device_name]["Interface"][interface_name]["IpAddress"] = str(addr.ip)
                        database[device_name]["Interface"][interface_name]["SubnetMask"] = str(addr.netmask)
                    #-------------------------------------------------------------------------------------
                    # Host/subnet mismatch detected, do not allow interface to generate the configuration
                    #-------------------------------------------------------------------------------------
                    else:
                        print ("  !-- invalid ip address, config not generated")
                        database[device_name]["Interface"][interface_name]["IpAddress"] = ""
                        database[device_name]["Interface"][interface_name]["SubnetMask"] = ""
                        CaptureError("Layer3",device_name,"Host/subnet mismatch",interface_name+":"+ip_address,"Not used")
                #--------------------------------------------------------------------------------------------
                # Wrong IP syntax used for the IP address, do not allow interface to generate configuration
                #--------------------------------------------------------------------------------------------
                except:
                    print ("  !-- invalid ip address, config not generated")
                    CaptureError("Layer3",device_name,"Invalid IP Syntax",interface_name+":"+ip_address,"Not used")
                    database[device_name]["Interface"][interface_name]["IpAddress"] = ""
                    database[device_name]["Interface"][interface_name]["SubnetMask"] = ""
        #-------------------------------------------------------------------------------
        # Check to see if port profile 1 is referencing a variable that actually exist
        #--------------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["PortProfile1"]:
            port_profile = database[device_name]["Interface"][interface_name]["PortProfile1"]
            if port_profile not in variables:
                print ("  !-- Invalid port profile 1 used for this port (%s)"%port_profile)
                database[device_name]["Interface"][interface_name]["PortProfile1"] = ""
                CaptureError("L2/L3",device_name,"Invalid/blank port profile 1",interface_name,"Not used")
        #-------------------------------------------------------------------------------
        # Check to see if port profile 2 is referencing a variable that actually exist
        #--------------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["PortProfile2"]:
            port_profile = database[device_name]["Interface"][interface_name]["PortProfile2"]
            if port_profile not in variables:
                print ("  !-- Invalid port profile 2 used for this port (%s)"%port_profile)
                database[device_name]["Interface"][interface_name]["PortProfile2"] = ""
                CaptureError("L2/L3",device_name,"Invalid/blank port profile 2",interface_name,"Not used")
        #------------------------------------------------------------------------------------------------
        # If the port is a trunk and allowing a range of vlans, make sure the vlans are created on switch
        #-------------------------------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["TrunkAllowedVlans"]:
            #-----------------------------------------------------
            # Check to see if there is a native vlan if it exists
            #-----------------------------------------------------
            if database[device_name]["Interface"][interface_name]["NativeVlan"]:
                native_vlan = int(database[device_name]["Interface"][interface_name]["NativeVlan"])
                if native_vlan not in database[device_name]["Vlans"]:
                    value = str(interface_name) + " (" + str(native_vlan) + ")"
                    print ("  !-- native vlan does not exists")
                    CaptureError("Layer2",device_name,"Invalid trunk native vlan",value,"Warning generated")

            allowed_vlans = database[device_name]["Interface"][interface_name]["TrunkAllowedVlans"]
            #------------------------------------------------------------------------------
            # Loop through all the vlans that are defined in the trunk allowed vlan column
            #------------------------------------------------------------------------------
            for vlan in allowed_vlans.split(","):
                #---------------------------------------------------------------------------------
                # If it detects a vlan range for this trunk then determine the start and end vlan
                #----------------------------------------------------------------------------------
                if "-" in vlan:
                    vlan_list = []
                    for multiple_vlans in vlan.split("-"):
                            vlan_list.append(multiple_vlans)
                    start_vlan = int(vlan_list[0])
                    end_vlan = int(vlan_list[1])
                    vlan_times_to_loop = end_vlan - start_vlan+1
                    vlan_in_loop = start_vlan-1
                    #---------------------------------------------
                    # Loop through all the vlans within the range
                    #---------------------------------------------
                    for x in range(vlan_times_to_loop):
                        vlan_in_loop = int(vlan_in_loop)+1
                        #---------------------------------------------------------------------------------------------
                        # Check each vlan individually to see whether it is in the vlan database for the device
                        #----------------------------------------------------------------------------------------------
                        if int(vlan_in_loop) not in database[device_name]["Vlans"]:
                            print ("  !-- Trunk allows VLAN %s which does not exist"%vlan_in_loop)
                            variable = interface_name + "(vlan: " + str(vlan_in_loop) + ")"
                            CaptureError("Layer2",device_name,"Trunk allowing invalid Vlan",variable,"Warning generated")
                #--------------------------------------------------------------------------------------------------
                # if no range is defined then check whether the single vlan is in the vlan database for the device
                #--------------------------------------------------------------------------------------------------
                else:
                    if int(vlan) not in database[device_name]["Vlans"]:
                        print ("  !-- Trunk allows VLAN %s which does not exist"%vlan)
                        CaptureError("Layer2",device_name,"Trunk using invalid Vlan",(interface_name+" ("+vlan+")"),"Warning generated")
        #-----------------------------------------------------------------------------------
        # Check to see if the data Vlan has been created on the switch via the "vlan" Tab
        #-----------------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["DataVlan"]:
            vlan = database[device_name]["Interface"][interface_name]["DataVlan"]
            if int(vlan) not in database[device_name]["Vlans"]:
                print ("  !-- data vlan not in vlan database")
                CaptureError("Layer2",device_name,"Data vlan does not exist",(interface_name+":vlan"+str(vlan)),"Warning generated")
        #-----------------------------------------------------------------------------------
        # Check to see if the voice Vlan has been created on the switch via the "vlan" Tab
        #-----------------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["VoiceVlan"]:
            vlan = database[device_name]["Interface"][interface_name]["VoiceVlan"]
            if int(vlan) not in database[device_name]["Vlans"]:
                print ("  !-- voice vlan not in vlan database")
                CaptureError("Layer2",device_name,"Voice vlan does not exist",(interface_name+":vlan"+str(vlan)),"Warning generated")
        #-----------------------------------------------------------------------------------
        # Check to see if the VRF has been created on the switch via the "vrf" Tab
        #-----------------------------------------------------------------------------------
        if database[device_name]["Interface"][interface_name]["VRF"]:
            vrf = database[device_name]["Interface"][interface_name]["VRF"]
            if vrf not in database[device_name]["VRF"]:
                print ("  !-- VRF has not been defined in the 'vrf' tab for this device")
                CaptureError("Layer3",device_name,"VRF does not exist",(interface_name+" ("+vrf+")"),"Warning generated")

    #-------------------------------------------------------------------
    # Generate the global configuration based on the profiles worksheet
    #-------------------------------------------------------------------
    def CreateGlobalConfig(self, device):
        first_time = True
        for profiles in database[device]["Profiles"]:
            if first_time:
                print ("!--------------------------")
                print ("! Global configuration")
                print ("!--------------------------")
                print ("hostname {}\n".format(device))
                first_time = False
            profile_name = profiles
            print ("! + Config from profile: %s"%profile_name)
            print ("%s"%variables[profile_name])
            print ("")

    #--------------------------------------------------------------
    # Generate the vlan configuration based on the vlan worksheet
    #--------------------------------------------------------------
    def CreateVlanConfig(self, device):
        first_time = True
        for vlan in sorted(database[device]["Vlans"]):
            if first_time:
                print ("!--------------------------")
                print ("! VLAN configuration")
                print ("!--------------------------")
                first_time = False
            vlan_name = database[device]["Vlans"][vlan]
            print ("vlan %s"%int(vlan))
            print ("  name %s"%vlan_name)
            print ("")
    #--------------------------------------------------------------
    # Generate the VRF configuration based on the vrf worksheet
    #--------------------------------------------------------------
    def CreateVrfConfig(self, device):
        first_time = True
        #-----------------------------------------------
        # Cycle through the VRF entries for the device
        #-----------------------------------------------
        for vrf in sorted(database[device]["VRF"]):
            if first_time:
                print ("!----------------------")
                print ("! VRF configuration")
                print ("!-----------------------")
                first_time = False
            #------------------------------------------------------------------------
            # VRF error checking has been shifted into the CheckVrfErrors function
            #------------------------------------------------------------------------
            self.CheckVrfErrors(device,vrf)
            #----------------------------------------
            # Generate the actual VRF configuration
            #----------------------------------------
            print ("ip vrf %s"%vrf)
            if database[device]["VRF"][vrf]["RD"]:
                print ("  rd %s"%database[device]["VRF"][vrf]["RD"])
            else:
                print ("  !- config cancelled, no RD defined")
            #-------------------------------------------------------------------------------
            # If there are multiple import route-targets, then create an entry for each one
            #-------------------------------------------------------------------------------
            for rt_import in database[device]["VRF"][vrf]["RTImport"].split(','):
                if rt_import:
                    if " " in rt_import:
                        rt_import = rt_import.replace(" ","")
                    print ("  route-target import %s"%rt_import)
            #-------------------------------------------------------------------------------
            # If there are multiple export route-targets, then create an entry for each one
            #-------------------------------------------------------------------------------
            for rt_export in database[device]["VRF"][vrf]["RTExport"].split(','):
                if rt_export:
                    if " " in rt_export:
                        rt_export = rt_export.replace(" ","")
                    print ("  route-target export %s"%rt_export)
            #---------------------------------------------------
            # Apply any profiles that are assigned to the VRF
            #---------------------------------------------------
            profile = database[device]["VRF"][vrf]["Profile"]
            if profile:
                print ("  %s"%variables[profile])

    #------------------------------------------------------------------
    # Generate the routing configuration based on the routing worksheet
    #------------------------------------------------------------------
    def CreateRoutingConfig(self, device):
        first_time = True
        #-----------------------------------------------
        # Cycle through the routing entries for the device
        #-----------------------------------------------
        for route in sorted(database[device]["Routes"]):
            if first_time:
                print ("!------------------------")
                print ("! Routing configuration")
                print ("!-------------------------")
                first_time = False
            #------------------------------------------------------------------------
            # Route error checking has been shifted into the CheckRoutingErrors function
            #------------------------------------------------------------------------
            self.CheckRoutingErrors(device,route)
            #----------------------------------------
            # Generate the actual routing configuration
            #----------------------------------------
            route_entry = database[device]["Routes"][route]["Route"]
            route_subnet = database[device]["Routes"][route]["SubnetMask"]
            route_vrf = database[device]["Routes"][route]["VRF"]
            route_nexthop = database[device]["Routes"][route]["NextHop"]
            route_name = database[device]["Routes"][route]["Name"]

            if route_entry:
                if route_vrf and route_subnet and route_nexthop and route_name:
                    print ("ip route vrf %s %s %s %s name %s"%(route_vrf,route_entry,route_subnet,route_nexthop,route_name))
                elif route_vrf and route_subnet and route_nexthop:
                    print ("ip route vrf %s %s %s %s"%(route_vrf,route_entry,route_subnet,route_nexthop))
                elif route_subnet and route_nexthop and route_name:
                    print ("ip route %s %s %s name %s"%(route_entry,route_subnet,route_nexthop,route_name))
                elif route_subnet and route_nexthop:
                    print ("ip route %s %s %s"%(route_entry,route_subnet,route_nexthop))


    #----------------------------------------------------------------------------
    # Generate the prefix-list configuration based on the prefix-list worksheet
    #----------------------------------------------------------------------------
    def CreatePrefixListConfiguration(self, device):
        first_time = True
        #-----------------------------------------------
        # Cycle through the routing entries for the device
        #-----------------------------------------------
        for prefix_name in sorted(database[device]["Prefix-Lists"]):
            if first_time:
                print ("!--------------------------")
                print ("! Prefix-list configuration")
                print ("!--------------------------")
                first_time = False
            #------------------------------------------------------------------------
            # Route error checking has been shifted into the CheckRoutingErrors function
            #------------------------------------------------------------------------
            self.CheckPrefixErrors(device,prefix_name)
            #----------------------------------------
            # Generate the actual routing configuration
            #----------------------------------------
            for seq in sorted(database[device]["Prefix-Lists"][prefix_name]):
                if database[device]["Prefix-Lists"][prefix_name][seq]["Valid"] == False:
                    print ("!-- error in {} seq [{}] entry not generated".format(prefix_name,seq))
                    continue
                prefix_entry = database[device]["Prefix-Lists"][prefix_name][seq]["Entry"]
                prefix_action = database[device]["Prefix-Lists"][prefix_name][seq]["Action"]
                print ("ip prefix-list {} seq {} {} {}".format(prefix_name,int(seq),prefix_action,prefix_entry))

    #-----------------------------------------------------------------------------------------
    # Generate the interface configuration based on layer2, layer3 and portchannel worksheets
    #-----------------------------------------------------------------------------------------
    def CreateInterfaceConfig(self, device,config_type):
        first_time = True
        #----------------------------------------------
        # Cycle through each interface for this device
        #----------------------------------------------
        for interface in sorted(database[device]["Interface"]):
            #----------------------------------------------------------------------------------
            # Do not show logical interfaces as physical config generation mode has been selected
            #----------------------------------------------------------------------------------
            if config_type == "Physical":
                if not database[device]["Interface"][interface]["PortMakeup"] == "Physical":
                    continue
            #----------------------------------------------------------------------------------
            # Do not show physical interfaces as logical config generation mode has been selected
            #----------------------------------------------------------------------------------
            elif config_type == "Logical":
                if not database[device]["Interface"][interface]["PortMakeup"] == "Logical":
                    continue

            if first_time and config_type == "Physical":
                print ("!-----------------------------------")
                print ("! Interface configuration (Physical)")
                print ("!-----------------------------------")
                first_time = False
            elif first_time and config_type == "Logical":
                print ("!-----------------------------------")
                print ("! Interface configuration (Logical)")
                print ("!-----------------------------------")
                first_time = False
            #------------------------------------------------------
            # Determine the interface type is Layer 2 or Layer 3
            #------------------------------------------------------
            interface_type = database[device]["Interface"][interface]["PortType"]
            #---------------------------------------------------------------------------------------------------
            # If it's a layer 3 port channel member interface, initalise the parent interface as no switchport
            #---------------------------------------------------------------------------------------------------
            if interface_type == "Layer3" and database[device]["Interface"][interface]["PortChannelParent"]:
                parent = database[device]["Interface"][interface]["PortChannelParent"]
                print ("!")
                print ("! + Layer 3 port channel member detected, set no switchport on logical interface:")
                print ("!-----------------")
                print ("interface %s"%parent)
                print ("  no switchport")
                print ("!-----------------")
            print ("interface %s"%interface)
            #-----------------------------------------------------------------------------------
            #  Interface error checking has been shifted into the CheckInterfaceErrors function
            #----------------------------------------------------------------------------------
            self.CheckInterfaceErrors(device,interface)
            #------------------------------------------------------
            # Generate Layer 2 configuration in the right sequence
            #------------------------------------------------------
            if interface_type == "Layer2":
                #-----------------------------
                # Define interface description
                #-----------------------------
                if database[device]["Interface"][interface]["Description"]:
                    value = database[device]["Interface"][interface]["Description"]
                    if u'\u2013' in value:
                        value = value.replace(u'\u2013', '-').encode('UTF-8')
                        value = value.decode('UTF-8')
                    print ("  description {}".format(value))
                #----------------------------------
                # Define the port as a switch port
                #----------------------------------
                print ("  switchport")
                #-------------------------
                # Define interface speed
                #-------------------------
                if database[device]["Interface"][interface]["Speed"]:
                    value = database[device]["Interface"][interface]["Speed"]
                    print ("  speed %s"%value)
                #-------------------------
                # Define interface duplex
                #-------------------------
                if database[device]["Interface"][interface]["Duplex"]:
                    value = database[device]["Interface"][interface]["Duplex"]
                    print ("  duplex %s"%value)
                #----------------------
                # Define interface MTU
                #----------------------
                if database[device]["Interface"][interface]["MTU"]:
                    value = database[device]["Interface"][interface]["MTU"]
                    print ("  mtu %s"%value)
                #---------------------------------------------
                # Define interface as a trunk and print vlans
                #---------------------------------------------
                if database[device]["Interface"][interface]["TrunkAllowedVlans"]:
                    value = database[device]["Interface"][interface]["TrunkAllowedVlans"]
                    print ("  switchport mode trunk")
                    print ("  switchport trunk allowed vlan {}".format(value))
                #------------------------------------------------------------------------------------------
                # Check to see if the interface has a native vlan configured, if so don't do anything yet
                #------------------------------------------------------------------------------------------
                if database[device]["Interface"][interface]["NativeVlan"]:
                    pass
                #-----------------------------------------------------------------------------------------------------------
                # If it doesn't have a native vlan, but is a member of a port channel, see if the logical interface has one
                #-----------------------------------------------------------------------------------------------------------
                elif database[device]["Interface"][interface]["PortChannelParent"]:
                    parent = database[device]["Interface"][interface]["PortChannelParent"]
                    #---------------------------------------------------------------------
                    # If it does, then inherit the native vlan from the logical interface
                    #---------------------------------------------------------------------
                    if database[device]["Interface"][parent]["NativeVlan"]:
                        native_vlan = int(database[device]["Interface"][parent]["NativeVlan"])
                        database[device]["Interface"][interface]["NativeVlan"] = int(native_vlan)
                        print ("  switchport mode trunk")
                #-----------------------------------------------------------------
                # Now actually print the native vlan if it the interface has one
                #-----------------------------------------------------------------
                if database[device]["Interface"][interface]["NativeVlan"]:
                    native_vlan = int(database[device]["Interface"][interface]["NativeVlan"])
                    print ("  switchport trunk native vlan {}".format(native_vlan))
                #-----------------------------------------
                # Define interface data vlan (access port)
                #-----------------------------------------
                if database[device]["Interface"][interface]["DataVlan"]:
                    value = database[device]["Interface"][interface]["DataVlan"]
                    print ("  switchport mode access")
                    print ("  switchport access vlan %s"%value)
                #------------------------------
                # Define interface voice vlan
                #------------------------------
                if database[device]["Interface"][interface]["VoiceVlan"]:
                    value = database[device]["Interface"][interface]["VoiceVlan"]
                    print ("  switchport voice vlan %s"%value)
                #-------------------------------------------
                # Generate the port profile 1 configuration
                #-------------------------------------------
                if database[device]["Interface"][interface]["PortProfile1"]:
                    value = database[device]["Interface"][interface]["PortProfile1"]
                    print ("%s"%variables[value])
                #-------------------------------------------
                # Generate the port profile 1 configuration
                #-------------------------------------------
                if database[device]["Interface"][interface]["PortProfile2"]:
                    value = database[device]["Interface"][interface]["PortProfile2"]
                    print ("%s"%variables[value])
                #----------------------------------------------------------------------------------------------
                # If it's a port-channel member interface, then configure channel-group on physical interface
                #-----------------------------------------------------------------------------------------------
                if database[device]["Interface"][interface]["PortChannelGroup"]:
                    value1 = int(database[device]["Interface"][interface]["PortChannelGroup"])
                    value2 = str(database[device]["Interface"][interface]["PortChannelMode"])
                    print ("  channel-group %s mode %s"%(value1,value2))
                #----------------------------------------------------
                # Check to see whether the port is activated or not
                #----------------------------------------------------
                if database[device]["Interface"][interface]["PortEnabled"]:
                    value = database[device]["Interface"][interface]["PortEnabled"]
                    if value in ("yes","Yes","no shutdown","No Shutdown"):
                        value = "no shutdown"
                    else:
                        value = "shutdown"
                    print ("  %s"%value)
                print ("!")
            #------------------------------------------------------
            # Generate Layer 3 configuration in the right sequence
            #------------------------------------------------------
            if interface_type == "Layer3":
                #-----------------------------
                # Define interface description
                #-----------------------------
                if database[device]["Interface"][interface]["Description"]:
                    value = database[device]["Interface"][interface]["Description"]
                    if u'\u2013' in value:
                        value = value.replace(u'\u2013', '-').encode('UTF-8')
                        value = value.decode('UTF-8')
                    print ("  description {}".format(value))

                #------------------------
                # Define interface VRF
                #------------------------
                if database[device]["Interface"][interface]["VRF"]:
                    value = database[device]["Interface"][interface]["VRF"]
                    print ("  ip vrf forwarding %s"%value)
                #-----------------------------
                # Define interface IP address
                #-----------------------------
                if database[device]["Interface"][interface]["IpAddress"]:
                    ip_address = database[device]["Interface"][interface]["IpAddress"]
                    subnet_mask = database[device]["Interface"][interface]["SubnetMask"]
                    print ("  ip address %s %s"%(ip_address,subnet_mask))
                #-----------------------
                # Define interface MTU
                #-----------------------
                if database[device]["Interface"][interface]["MTU"]:
                    value = database[device]["Interface"][interface]["MTU"]
                    print ("  mtu %s"%value)
                #-------------------------------------------
                # Generate the port profile 1 configuration
                #-------------------------------------------
                if database[device]["Interface"][interface]["PortProfile1"]:
                    value = database[device]["Interface"][interface]["PortProfile1"]
                    print ("%s"%variables[value])
                #-------------------------------------------
                # Generate the port profile 1 configuration
                #-------------------------------------------
                if database[device]["Interface"][interface]["PortProfile2"]:
                    value = database[device]["Interface"][interface]["PortProfile2"]
                    print ("%s"%variables[value])
                #---------------------------------------------------------------------
                # If it's a port-channel interface, define the group number and type
                #---------------------------------------------------------------------
                if database[device]["Interface"][interface]["PortChannelGroup"]:
                    print ("  no switchport")
                    value1 = int(database[device]["Interface"][interface]["PortChannelGroup"])
                    value2 = str(database[device]["Interface"][interface]["PortChannelMode"])
                    print ("  channel-group %s mode %s"%(value1,value2))
                #----------------------------------------------------
                # Check to see whether the port is activated or not
                #----------------------------------------------------
                if database[device]["Interface"][interface]["PortEnabled"]:
                    value = database[device]["Interface"][interface]["PortEnabled"]
                    if value in ("yes","Yes","no shutdown","No Shutdown"):
                        value = "no shutdown"
                    else:
                        value = "shutdown"
                    print ("  %s"%value)
                    print ("!")

    #-------------------------------------------------------------------------------------
    # This function will check the error database and notify the operator of any problems
    #-------------------------------------------------------------------------------------
    def CreateErrorReport(self):
        if (errors):
            os.system('cls' if os.name == 'nt' else 'clear')

            print ("!----------------------------------------------------------------------------")
            print ("! The following errors have been detected and should be addressed:")
            print ("!----------------------------------------------------------------------------")
            print("============================================================================================================")
            print(" [%-10s] [%-15s] [%-30s] [%-20s] [%-15s]  "%("Location","Device","ErrorMSg","Variable","Action"))
            print("============================================================================================================")

            for name in sorted(errors):
                for no in errors[name]:
                    location = errors[name][no]["Location"]
                    error_msg = errors[name][no]["ErrorMsg"]
                    error_variable = errors[name][no]["Variable"]
                    action = errors[name][no]["Action"]
                    print (" %-10s   %-15s   %-30s   %-20s   %-15s"%(location,name,error_msg,error_variable,action))
        else:
            print ("No errors detected during configuration build.")

    #------------------------------------------------------------
    # This function will generate all the relevant configuration
    #------------------------------------------------------------
    def CreateAllConfig(self):
        console = sys.__stdout__
        for device in sorted(database):
            devices.append(device)
            #today = datetime.date.today()
            sys.stdout = Logger("%s-config.txt"%device)
            print ("********************************************")
            print ("!    Device configuration for %s"%device)
            print ("********************************************")
            self.CreateGlobalConfig(device)
            self.CreateVlanConfig(device)
            self.CreateVrfConfig(device)
            self.CreateInterfaceConfig(device,"Physical")
            self.CreateInterfaceConfig(device,"Logical")
            self.CreateRoutingConfig(device)
            self.CreatePrefixListConfiguration(device)
            sys.stdout = console

#---------------
# Show the menu
#---------------
def ShowMenu():
    console = sys.__stdout__

    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print ("============================================================")
        print ("Cisco Config Generator %s"%__version__)
        print ("============================================================")
        path = os.path.join(os.path.dirname(__file__))
        #path = os.path.dirname(os.path.realpath(__file__))
        print ("Build file used  : [%s]"%filename)
        print ("Config path used : [%s]"%path)
        print ("")
        print ("\t1. Generate configuration")
        print ("\t2. Quit")
        print ("============================================================")
        selection=input("Please Select: ")
        print ("============================================================")
        #-----------------------------------------------------------
        # Perform action based on what key is selected by the user
        #-----------------------------------------------------------
        if selection =='1':
            print ("Now generating configuration files.....")
            #----------------------------------
            # Start saving output to debug.log
            #----------------------------------
            today = datetime.date.today()
            sys.stdout = Logger("debug.log")
            #-------------------------------------------
            # Read all data from the build spreadsheet
            #-------------------------------------------
            record = ReadConfig()
            record.ReadVariables()
            record.ReadProfiles()
            record.ReadPortChannel()
            record.ReadVlans()
            record.ReadVrf()
            record.ReadLayer2()
            record.ReadLayer3()
            record.ReadRouting()
            record.ReadPrefixList()
            sys.stdout = console
            #----------------------------------------------------------------------------------------
            # Start generating configuration files (unique txt file will be generated for each device)
            #-----------------------------------------------------------------------------------------
            config = GenerateConfig()
            config.CreateAllConfig()
            #----------------------------------
            # Start saving output to errors.log
            #----------------------------------
            sys.stdout = Logger("errors.log")
            config.CreateErrorReport()
            sys.stdout = console
            #---------------------------------------------------------------------
            # Show the final screen which list the files which have been created
            #---------------------------------------------------------------------
            os.system('cls' if os.name == 'nt' else 'clear')
            today = datetime.date.today()
            print ("============================================================")
            print ("Cisco Config Generator %s"%__version__)
            print ("============================================================")
            print ("The following files have been generated:\n")
            print ("  --debug.log")
            print ("  --errors.log")
            for device in sorted(devices):
                if device in errors:
                    print ("  --%s-config.txt    [%-10s]"%(device,"errors found"))
                else:
                    print ("  --%s-config.txt    [%-10s]"%(device,"successful"))
            return

        elif selection == '2':
            print ("Program closed.")
            return
        else:
            os.system('cls' if os.name == 'nt' else 'clear')
            print ("Unknown Option Selected.")



def main(argv):
    arg_length = len(sys.argv)
    if arg_length < 2:
        print ("============================================================")
        print ("Cisco Config Generator %s"%__version__)
        print ("============================================================")
        print ("Usage: %s <filename.xls>"%sys.argv[0])
        exit()

    if sys.argv[1]:
        global filename
        filename = sys.argv[1]
    try:
        workbook = xlrd.open_workbook(filename)
        ShowMenu()
    except IOError:
        print ("Unable to open: %s"% sys.argv[1])
        print ("Program aborted.")
        exit()


if __name__ == '__main__':
    main(sys.argv)

