import xlrd
import re
import sys
from netaddr import *
from operator import itemgetter

__author__ = 'Abdul Karim El-Assaad'
__version__ = '(CCG) Version: 2.1 BETA (25/9/2014)'  # Cisco Config Generator version

error_db = {}

raw_db = {}             # Stores all the raw data from the build spreadsheet
device_list = []        # Stores the valid device list from raw_db
worksheet_list = []     # Stores the names of all worksheets in the build spreadsheet
column_list = {}

# Create a unique global dictionary for each worksheet
config_templates = {}   # Stores the config templates from raw_db
variable_list = {}      # Stores the valid variables from raw_db
profile_list = {}       # Stores the profiles from raw_db
vlan_list = {}          # Stores the valid vlans from raw_db
vrf_list = {}           # Stores the valid vrfs from raw_db
interface_list = {}        # Stores the layer2 information from raw_db
static_route_list = {}  # Stores the routing information from raw_db
prefix_list = {}        # Stores the prefix-list information from raw_db
portchannel_list = {}   # Stores the port-channel list information from raw_db

global filename

'''
=============================================================================
How to add a new worksheet:
    * Create a new global dictionary or list (see above)
    * Update RemoveEmptyRowsFromDB() function to include the new required columns
    * Create a function called Get<whatever>List
    * Create a function called Create<whatever>Configuration
=============================================================================
'''

#-----------------------------------------
# Used to redirect output to a text file
#-----------------------------------------
class Logger(object):
    def __init__(self, filename="Default.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "w")
    def write(self, message):
        #self.terminal.write(message)   # Shows output to screen
        self.log.write(message)         # Writes output to file

#-----------------------------------------
# Main class which has all the functions
#-----------------------------------------
class Config(object):
    def __init__(self):
        self.CreateRawDb()

    # ----------------------------------------------------------------------
    # Read the content of the build spreadsheet into the raw_db dictionary
    # To access call:  raw_db["worksheet_name"][row_number]["Column name"]
    # ----------------------------------------------------------------------
    def CreateRawDb(self):
        wb = xlrd.open_workbook(filename)
        global raw_db
        global worksheet_list
        global device_list
        global column_list
        global error_db

        temp_db = []

        for i, worksheet in enumerate(wb.sheets()):
            if worksheet.name == "Instructions":
                continue
            header_cells = worksheet.row(0)
            num_rows = worksheet.nrows - 1
            curr_row = 0
            header = [each.value for each in header_cells]
            column_list[worksheet.name] = header
            # Add a column that doesn't exist in the worksheet
            header.append("Row")
            #-------------------------------------------------------------------------------------
            # Iterate over each row in each worksheet and store the info in the raw_db dictionary
            #-------------------------------------------------------------------------------------
            while curr_row < num_rows:
                curr_row += 1
                row = [int(each.value) if isinstance(each.value, float)
                       else each.value
                       for each in worksheet.row(curr_row)]
                # Add the row number to each record
                row.append(curr_row+1)
                value_dict = dict(zip(header, row))
                temp_db.append(value_dict)
            else:
                #print ("raw_db: added '{}'".format(worksheet.name))
                raw_db[worksheet.name] = temp_db
                #---------------------------------------------------------------------
                # Grab all the unique device worksheet names in the build spreadsheet
                #---------------------------------------------------------------------
                worksheet_list.append(worksheet.name)
                error_db[worksheet.name] = []

                #------------------------------------------------------------------------
                # Re-initalise the temp_db database so it's ready for the next worksheet
                #------------------------------------------------------------------------
                temp_db = []

#-------------------------------------------------------------------------------------
# Read through raw_db and start storing relevant information in their global database
#-------------------------------------------------------------------------------------

    #-----------------------------------------------------------------------
    # Create a database which contains a list devices from every worksheet
    #------------------------------------------------------------------------
    def GetDeviceList(self):
        global device_list
        raw_devices = []
        for worksheet in worksheet_list:
            for row,entry in enumerate(raw_db[worksheet]):
                if raw_db[worksheet][row].get("Device Name"):
                    if "!" in raw_db[worksheet][row].get("Device Name"):
                        continue
                    value = raw_db[worksheet][row].get("Device Name","none")
                    value = value.strip()
                    if value not in raw_devices:
                        raw_devices.append(value)
        device_list = sorted(raw_devices)


    #-----------------------------------------------------------
    # Create a database that contains the VLANs for each device
    #-----------------------------------------------------------
    def GetVlanList(self):
        global vlan_list
        for row_no in range(len(raw_db["vlans"])):
            device_name = raw_db["vlans"][row_no]["Device Name"]
            if vlan_list.get(device_name):
                continue
            else:
                vlan_list[device_name] = {}
        for row_no in range(len(raw_db["vlans"])):
            device_name = raw_db["vlans"][row_no]["Device Name"]
            vlan_no = raw_db["vlans"][row_no]["VLAN No"]
            vlan_name = raw_db["vlans"][row_no]["VLAN Name"]
            vlan_list[device_name][vlan_no] = vlan_name

    #-----------------------------------------------------------
    # Create a database that contains the VRFs for each device
    #-----------------------------------------------------------
    def GetVrfList(self):
        global vrf_list
        for row_no in range(len(raw_db["vrf"])):
            vrf_columns = {key: [] for key in column_list["vrf"]}
            device_name = raw_db["vrf"][row_no]["Device Name"]
            vrf = raw_db["vrf"][row_no]["VRF"]
            if not vrf_list.get(device_name):
                vrf_list[device_name] = {}
            if not vrf_list[device_name].get(vrf):
                vrf_list[device_name][vrf] = vrf_columns
            for column in vrf_columns:
                vrf_list[device_name][vrf][column] = raw_db["vrf"][row_no][column]
            # If there are multiple route-targets grab them all
            current_import = vrf_list[device_name][vrf]["Import RT  (separated by commas)"]
            current_import = current_import.strip()
            current_import = current_import.replace(" ","")
            current_export = vrf_list[device_name][vrf]["Export RT  (separated by commas)"]
            current_export = current_export.strip()
            current_export = current_export.replace(" ","")
            new_import = current_import.split(",")
            new_export = current_export.split(",")
            vrf_list[device_name][vrf]["Import RT  (separated by commas)"] = new_import
            vrf_list[device_name][vrf]["Export RT  (separated by commas)"] = new_export

    #------------------------------------------------------------------
    # Create a database that contains the interfaces for each device
    #------------------------------------------------------------------
    def GetInterfaceList(self):
        global interface_list
        for row_no in range(len(raw_db["interfaces"])):
            interface_columns = {key: [] for key in column_list["interfaces"]}
            device_name = raw_db["interfaces"][row_no]["Device Name"]
            port = raw_db["interfaces"][row_no]["Interface"]
            if not interface_list.get(device_name):
                interface_list[device_name] = {}
            if not interface_list[device_name].get(port):
                interface_list[device_name][port] = interface_columns
            for column in interface_columns:
                interface_list[device_name][port][column] = raw_db["interfaces"][row_no][column]
            self.AddInterface(device_name,port,True)

    #-------------------------------------------------------------------
    # Create a database that contains the static routes for each device
    #-------------------------------------------------------------------
    def GetStaticRouteList(self):
        global static_route_list
        for row_no in range(len(raw_db["static routes"])):
            static_route_columns = {key: [] for key in column_list["static routes"]}
            device_name = raw_db["static routes"][row_no]["Device Name"]
            route = raw_db["static routes"][row_no]["Route (x.x.x.x/x)"]
            if not static_route_list.get(device_name):
                static_route_list[device_name] = {}
            if not static_route_list[device_name].get(route):
                static_route_list[device_name][route] = static_route_columns
            for column in static_route_columns:
                static_route_list[device_name][route][column] = raw_db["static routes"][row_no][column]

            new_route = IPNetwork(route)
            static_route_list[device_name][route]["Route"] = str(new_route.ip)
            static_route_list[device_name][route]["Subnet"] = str(new_route.netmask)


    def GetPrefixList(self):
        global prefix_list
        for row_no in range(len(raw_db["prefix-list"])):
            device_name = raw_db["prefix-list"][row_no]["Device Name"]
            prefix_name = raw_db["prefix-list"][row_no]["Prefix-List Name"]
            prefix_seq  = raw_db["prefix-list"][row_no]["Prefix-List Sequence No"]
            prefix_action  = raw_db["prefix-list"][row_no]["Prefix-List Action (permit/deny)"]
            prefix_entry  = raw_db["prefix-list"][row_no]["Prefix-List Entry"]
            if not prefix_list.get(device_name):
                prefix_list[device_name] = {}
            if not prefix_list[device_name].get(prefix_name):
                temp_list = []
                prefix_list[device_name][prefix_name] = temp_list
                temp_db = {"sequence":"","action":"","entry":""}
                prefix_list[device_name][prefix_name].append(temp_db)
                prefix_list[device_name][prefix_name][0]["sequence"] = prefix_seq
                prefix_list[device_name][prefix_name][0]["action"] = prefix_action
                prefix_list[device_name][prefix_name][0]["entry"] = prefix_entry
            elif not self.GetPrefixSeqNo(device_name,prefix_name,prefix_seq):
                list_no = len(prefix_list[device_name][prefix_name])
                temp_db = {"sequence":"","action":"","entry":""}
                prefix_list[device_name][prefix_name].append(temp_db)
                prefix_list[device_name][prefix_name][list_no]["sequence"] = prefix_seq
                prefix_list[device_name][prefix_name][list_no]["action"] = prefix_action
                prefix_list[device_name][prefix_name][list_no]["entry"] = prefix_entry


    def GetPortChannelList(self):
        global portchannel_list
        for row_no in range(len(raw_db["portchannels"])):
            portchannel_column = {key: [] for key in column_list["portchannels"]}
            device_name = raw_db["portchannels"][row_no]["Device Name"]
            interface = raw_db["portchannels"][row_no]["Interface"]
            if not portchannel_list.get(device_name):
                portchannel_list[device_name] = {}
            if not portchannel_list[device_name].get(interface):
                portchannel_list[device_name][interface] = portchannel_column
            for column in portchannel_column:
                portchannel_list[device_name][interface][column] = raw_db["portchannels"][row_no][column]
        self.UpdatePortChannels()


    # Will iterate over the portchannel_list and try to add a new interface for each parent
    # or member interfaces.

    def UpdatePortChannels(self):
        for device_name in sorted(portchannel_list):
            for interface in sorted(portchannel_list[device_name]):

                pc_enabled = portchannel_list[device_name][interface]["Interface Enabled (yes/no)"]
                pc_group = portchannel_list[device_name][interface]["Port-Channel Group"]
                pc_mode = portchannel_list[device_name][interface]["Port-Channel Mode (active/on/etc)"]
                pc_type = portchannel_list[device_name][interface]["Port-Channel Type (layer2 or layer3)"]
                pc_members = portchannel_list[device_name][interface]["Port-Channel Members (separated by commas)"]
                pc_description = portchannel_list[device_name][interface]["Description"]

                self.AddInterface(device_name,interface)
                if pc_enabled:
                    interface_list[device_name][interface]["Interface Enabled (yes/no)"] = pc_enabled
                if pc_description:
                    interface_list[device_name][interface]["Description"] = pc_description

                for member in pc_members.split(","):
                    member = member.strip()
                    self.AddInterface(device_name,member)
                    interface_list[device_name][member]["PC-Group"] = pc_group
                    interface_list[device_name][member]["PC-Mode"] = pc_mode
                    interface_list[device_name][member]["PC-Type"] = pc_type
                    interface_list[device_name][interface]["PC-Members"].append(member)

    def GetVariableList(self):
        global variable_list
        for row_no in range(len(raw_db["variables"])):
            variable_name = raw_db["variables"][row_no]["Variable"]
            variable_value = raw_db["variables"][row_no]["Variable Value"]
            if "+" in variable_name:
                continue
            elif variable_name in variable_list:
                continue
            variable_list[variable_name] = variable_value

    def GetConfigTemplateList(self):
        global config_templates
        temp_list = []
        for row_no in range(len(raw_db["config-templates"])):
            line = raw_db["config-templates"][row_no]["Enter config templates below this line:"]
            match = re.search(r'Config Template: \[(.*?)\]', line,re.IGNORECASE)
            if not line:
                continue
            if match:
                temp_list = []
                config_templates[match.group(1)] = temp_list
                continue
            temp_list.append(line)

        # Cycle through each line of the config template to update dynamic variables
        for entry in config_templates:
            temp_list = []
            for line in config_templates[entry]:
                match = re.findall(r'\[(.*?)\]', line)
                valid_variable = False
                if (match):
                    for variable_name in match:
                        lookup_variable = self.GetVariable(variable_name)
                        if lookup_variable:
                            valid_variable = True
                            line = line.replace(variable_name,lookup_variable)
                        else:
                            error_db["config-templates"].append("Config-Template: '{}' referenced embedded variable '{}' which does not exist".format(entry,variable_name))
                            continue
                    if (valid_variable):
                        line = line.replace("[","")
                        line = line.replace("]","")
                temp_list.append(line)
            config_templates[entry] = temp_list


    def GetProfileList(self):
        global profile_list
        global error_db
        for row_no in range(len(raw_db["profiles"])):
            device_name = raw_db["profiles"][row_no]["Device Name"]
            variable = raw_db["profiles"][row_no]["Template or Variable"]
            position = raw_db["profiles"][row_no]["Position (Default: Start)"]
            if not position:
                position = "Start"
            if not self.GetVariable(variable):
                if not self.GetConfigTemplate(variable):
                    error_db["profiles"].append("Row ({}): Device '{}' referenced variable '{}' which does not exist".format(raw_db["profiles"][row_no]["Row"],device_name,variable))
                    continue

#            if device_name not in profile_list:
#                temp_list = []
#                profile_list[device_name] = temp_list
#            profile_list[device_name].append(variable)

            if device_name not in profile_list:
                profile_temp = []
                position_temp = []
                profile_list[device_name] = {}
                profile_list[device_name]["Profile"] = profile_temp
                profile_list[device_name]["Position"] = position_temp
            profile_list[device_name]["Profile"].append(variable)
            profile_list[device_name]["Position"].append(position)


    #-------------------------------------------
    # Manual manipulation of one of the lists
    #-------------------------------------------
    ''' manual_entries will result in manual entries always being triggered.
        This will be created when:
        a) GetInterfaceList is called (i.e. reading the info from the interfaces tab
        b) UpdatePortChannels is called (i.e. only if new interface is detected)
    '''
    def AddInterface(self,device_name,interface,manual_entries=False):
        global interface_list

        interface_columns = {key: [] for key in column_list["interfaces"]}
        if not interface_list.get(device_name):
            interface_list[device_name] = {}
        if not interface_list[device_name].get(interface):
            manual_entries = True
            interface_list[device_name][interface] = interface_columns

        # Any manual entries that are not directly from the spreadsheet below
        if manual_entries:
            error_list = []
            member_list = []
            interface_list[device_name][interface]["Errors"] = error_list
            interface_list[device_name][interface]["Interface"] = interface
            interface_list[device_name][interface]["PC-Group"] = ""
            interface_list[device_name][interface]["PC-Mode"] = ""
            interface_list[device_name][interface]["PC-Type"] = ""
            interface_list[device_name][interface]["PC-Members"] = member_list
            interface_list[device_name][interface]["PC-Parent"] = ""

    # ---------------------------------------------------
    # Get specific values from their respective database
    # ---------------------------------------------------
    def GetVariable(self, variable_name):
        if variable_list.get(variable_name):
            return variable_list[variable_name]

    def GetVlan(self, device_to_find, vlan_to_find):
        if vlan_list[device_to_find].get(vlan_to_find):
            return vlan_list[device_to_find][vlan_to_find]

    def GetConfigTemplate(self, template_name):
        if config_templates.get(template_name):
            return config_templates[template_name]

    def GetPrefixSeqNo(self, device_name,prefix_name,sequence_number):
        global prefix_list
        if not prefix_list.get(device_name):
            return
        if not prefix_list[device_name].get(prefix_name):
            return
        for row, entry in enumerate(prefix_list[device_name][prefix_name]):
            if prefix_list[device_name][prefix_name][row]["sequence"] == sequence_number:
                return prefix_list[device_name][prefix_name][row]

    def GetIP(self,IP_address,mode="IOS"):
        if not IP_address:
            return
        if "/" not in IP_address:
            # Generate error message
            return IP_address
        if mode == "NXOS":
            string = ("{}".format(IP_address))
            return string
        elif mode == "IOS":
            full_address = IPNetwork(IP_address)
            string = ("{} {}".format(full_address.ip,full_address.netmask))
            return string

    def GetInterfaceType(self,interface):
        logical_interfaces = ["Po","Tu","Lo","Vl"]
        for type in logical_interfaces:
            if type in interface:
                return "Logical"
        return "Physical"

    def GetTrunkVlans(self,device,interface):
        if not interface_list[device][interface].get("Trunk Allowed VLANs (separated by commas)"):
            return
        allowed_vlans_raw = interface_list[device][interface]["Trunk Allowed VLANs (separated by commas)"]
        allowed_vlans_raw = allowed_vlans_raw.replace(" ","")
        allowed_vlans_raw = allowed_vlans_raw.strip()
        allowed_vlans_raw = allowed_vlans_raw.split(",")
        allowed_vlans_new = []

        for vlan in allowed_vlans_raw:
            if "-" not in vlan:
                allowed_vlans_new.append(vlan)
            elif "-" in vlan:
                vlan_range = []
                for entry in vlan.split("-"):
                    vlan_range.append(entry)
                vlan_range_start = int(vlan_range[0])
                vlan_range_end = int(vlan_range[1])

                while vlan_range_start <= vlan_range_end:
                    allowed_vlans_new.append(vlan_range_start)
                    vlan_range_start +=1
        return allowed_vlans_new

    # ---------------------------------------------
    # Check the interface for specific conditions
    # ---------------------------------------------
    def is_switch_port(self,device,interface):
        if interface_list[device][interface]["Data VLAN"]:
            return True
        elif interface_list[device][interface]["Voice VLAN"]:
            return True
        elif interface_list[device][interface]["Trunk Allowed VLANs (separated by commas)"]:
            return True
        elif "layer2" in interface_list[device][interface]["PC-Type"]:
            return True

    def is_routed_port(self,device,interface):
        if "Logical" in self.GetInterfaceType(interface):
            return False
        elif interface_list[device][interface]["IP Address (x.x.x.x/x)"]:
            return True
        elif "layer3" in interface_list[device][interface]["PC-Type"]:
            return True

    def is_trunk_port(self,device,interface):
        if interface_list[device][interface].get("Trunk Allowed VLANs (separated by commas)"):
            return True

    def is_data_port(self,device,interface):
        if interface_list[device][interface].get("Data VLAN"):
            return True

    def is_voice_port(self,device,interface):
        if interface_list[device][interface].get("Voice VLAN"):
            return True

    def is_interface_enabled(self,device,interface):
        if interface_list[device][interface]["Interface Enabled (yes/no)"]:
            if "yes" in interface_list[device][interface]["Interface Enabled (yes/no)"]:
                return True
            if "Yes" in interface_list[device][interface]["Interface Enabled (yes/no)"]:
                return True

    def is_portchannel_member(self,device,interface):
        if interface_list[device][interface]["PC-Group"]:
            return True

    def is_portchannel_parent(self,device,interface):
        if interface_list[device][interface]["PC-Members"]:
            return True

    def is_valid_portchannel(self,device,interface):
        if portchannel_list[device].get(interface):
            if interface_list[device][interface]:
                return True

    def is_valid_variable(self,variable_name):
        if variable_list.get(variable_name):
            return True

    def is_valid_vlan(self,device,vlan):
        if vlan_list.get(device):
            if vlan_list[device].get(int(vlan)):
                return True

    def is_valid_trunk(self,device,interface):
        if not self.is_trunk_port(device,interface):
            return False
        trunk_vlans = self.GetTrunkVlans(device,interface)
        valid_trunk = True
        for vlan in trunk_vlans:
            if not self.is_valid_vlan(device,vlan):
                valid_trunk = False
        if valid_trunk:
            return True

    def is_valid_vrf(self,device,vrf):
        if vrf_list.get(device):
            if vrf_list[device].get(vrf):
                return True

    def is_valid_ipaddress(self,ip_address_to_check):
        ip_address = IPNetwork(ip_address_to_check)
        if IPAddress(ip_address.ip) in IPNetwork(ip_address).iter_hosts():
            return True

    def has_variable1_configured(self,device,interface):
        if interface_list[device][interface].get("Variable 1"):
            return True

    def has_variable2_configured(self,device,interface):
        if interface_list[device][interface].get("Variable 2"):
            return True

    def has_description_configured(self,device,interface):
        if interface_list[device][interface]["Description"]:
            return True

    def has_mtu_configured(self,device,interface):
        if interface_list[device][interface]["MTU"]:
            return True

    def has_vrf_configured(self,device,interface):
        if interface_list[device][interface]["VRF (leave blank if global)"]:
            return True

    def has_ip_configured(self,device,interface):
        if interface_list[device][interface]["IP Address (x.x.x.x/x)"]:
            return True

    def has_nativevlan_configured(self,device,interface):
        if interface_list[device][interface]["Trunk Native VLAN"]:
            return True

    def has_speed_configured(self,device,interface):
        if interface_list[device][interface]["Speed"]:
            return True

    def has_duplex_configured(self,device,interface):
        if interface_list[device][interface]["Duplex"]:
            return True

    def has_importrt_configured(self,device,vrf):
        if vrf_list[device][vrf].get("Import RT  (separated by commas)"):
            return True

    def has_exportrt_configured(self,device,vrf):
        if vrf_list[device][vrf].get("Export RT  (separated by commas)"):
            return True

    def has_rd_configured(self,device,vrf):
        if vrf_list[device][vrf].get("RD"):
            return True



    # -------------------------------------------------------------------------------------
    # Check each worksheet/row to make sure that the required information is present.
    # Otherwise the row will be removed from the database as it will be considered invalid
    # -------------------------------------------------------------------------------------
    def RemoveEmptyRowsFromDB(self):

        db_required_columns = {key: [] for key in worksheet_list}
        #--------------------------------------------------------------------
        # If the worksheet needs to have a valid column entry, define it below
        #--------------------------------------------------------------------
        db_required_columns["variables"] = ["Variable", "Variable Value"]
        db_required_columns["profiles"]= ["Device Name","Template or Variable"]
        db_required_columns["vrf"] = ["Device Name","VRF","RD"]
        db_required_columns["vlans"]= ["Device Name","VLAN No","VLAN Name"]
        db_required_columns["interfaces"]= ["Device Name","Interface"]
        db_required_columns["static routes"]= ["Device Name","Route (x.x.x.x/x)","Next Hop"]
        db_required_columns["portchannels"]= ["Device Name","Interface","Port-Channel Group","Port-Channel Mode (active/on/etc)","Port-Channel Type (layer2 or layer3)","Port-Channel Members (separated by commas)"]

        #--------------------------------------------
        # Search for invalid rows and update database
        #--------------------------------------------
        console = sys.__stdout__
        sys.stdout = Logger("ccg-ignored.txt")
        for worksheet in worksheet_list:
            for entry in db_required_columns[worksheet]:
                temp_db = []
                for row_no in range(len(raw_db[worksheet])):
                    if not raw_db[worksheet][row_no][entry]:

                        print ("[{}]-row:{} has empty cell value for column: {}  (IGNORED)".format(worksheet,raw_db[worksheet][row_no]["Row"],entry))
                        continue
                    if "$" in str(raw_db[worksheet][row_no][entry]):
                        continue
                    temp_db.append(raw_db[worksheet][row_no])
                else:
                    raw_db[worksheet] = temp_db
        sys.stdout = console

    #-------------------------------------------------------------
    # Functions to actually show the output of the configuration
    #-------------------------------------------------------------

    def CreateGlobalConfig(self,device_name,config_position="Start"):
        if not profile_list.get(device_name):
            return

        print ("!---------------------------------")
        print ("! Global configuration ({}) ".format(config_position))
        print ("!---------------------------------")
        for number, profile in enumerate(profile_list[device_name]["Profile"]):
            profile_name = profile_list[device_name]["Profile"][number]
            profile_position_type = profile_list[device_name]["Position"][number]
            if config_position not in profile_position_type:
                continue
            if self.GetConfigTemplate(profile_name):
                print ("\n! [{}]:".format(profile_name))
                for line in self.GetConfigTemplate(profile_name):
                    print ("{}".format(line))
            elif self.GetVariable(profile_name):
                print ("\n! [{}]:".format(profile_name))
                print ("{}".format(self.GetVariable(profile_name)))

    def CreateVlanConfig(self,device_name):
        if not vlan_list.get(device_name):
            return
        print ("!---------------------------------")
        print ("! VLAN configuration ")
        print ("!---------------------------------")
        for vlan in sorted(vlan_list[device_name]):
            print ("vlan {}".format(vlan))
            print (" name {}".format(vlan_list[device_name].get(vlan)))

    def CreateVrfConfig(self,device_name):
        if not vrf_list.get(device_name):
            return
        print ("!---------------------------------")
        print ("! VRF configuration ")
        print ("!---------------------------------")
        for vrf in sorted(vrf_list[device_name]):
            print ("ip vrf {}".format(vrf))
            if self.has_rd_configured(device_name,vrf):
                print ("  rd {}".format(vrf_list[device_name][vrf]["RD"]))
            if self.has_importrt_configured(device_name,vrf):
                for route_target in vrf_list[device_name][vrf]["Import RT  (separated by commas)"]:
                    print ("  route-target import {}".format(route_target))
            if self.has_exportrt_configured(device_name,vrf):
                for route_target in vrf_list[device_name][vrf]["Export RT  (separated by commas)"]:
                    print ("  route-target export {}".format(route_target))
            if vrf_list[device_name][vrf]["Variable"]:
                if self.is_valid_variable(vrf_list[device_name][vrf]["Variable"]):
                    print (self.GetVariable(vrf_list[device_name][vrf]["Variable"]))


    def CreateStaticRouteConfig(self,device_name):
        if not static_route_list.get(device_name):
            return
        print ("!---------------------------------")
        print ("! Static routing configuration ")
        print ("!---------------------------------")
        for route in sorted(static_route_list[device_name]):
            route_vrf = static_route_list[device_name][route]["VRF (leave blank if global)"]
            route_entry = static_route_list[device_name][route]["Route"]
            route_subnet = static_route_list[device_name][route]["Subnet"]
            route_nexthop = static_route_list[device_name][route]["Next Hop"]
            route_name = static_route_list[device_name][route]["Route Name (no spaces)"]

            if route_entry:
                if route_vrf and route_subnet and route_nexthop and route_name:
                    print ("ip route vrf {} {} {} name {}".format(route_vrf,route_entry,route_subnet,route_name))
                elif route_vrf and route_subnet and route_nexthop:
                    print ("ip route vrf {} {} {}".format(route_vrf,route_entry,route_subnet))
                elif route_subnet and route_nexthop and route_name:
                    print ("ip route {} {} {} name {}".format(route_entry,route_subnet,route_nexthop,route_name))
                elif route_subnet and route_nexthop:
                    print ("ip route {} {} {}".format(route_entry,route_subnet,route_nexthop))

    def CreatePrefixConfig(self,device_name):
        if not prefix_list.get(device_name):
            return
        print ("!---------------------------------")
        print ("! Prefix-list configuration ")
        print ("!---------------------------------")

        for prefix_name in sorted(prefix_list[device_name]):
            oldlist = prefix_list[device_name][prefix_name]
            newlist = sorted(oldlist, key=itemgetter('sequence'))
            for entry in newlist:
                pl_sequence = entry["sequence"]
                pl_action = entry["action"]
                pl_entry = entry["entry"]
                print ("ip prefix-list {} seq {} {} {}".format(prefix_name,pl_sequence,pl_action,pl_entry))
            else:
                print ("!")

    def CreateInterfaceConfig(self,device_name,config_mode):
        if not interface_list.get(device_name):
            return
        first_match = True
        for interface in sorted(interface_list[device_name]):
            if config_mode == "Physical":
                if not self.GetInterfaceType(interface) == "Physical":
                    continue
            elif config_mode == "Logical":
                if not self.GetInterfaceType(interface) == "Logical":
                    continue
            if (first_match and "Physical" in config_mode):
                print ("!--------------------------------------------")
                print ("! Interface configuration (Physical) ")
                print ("!--------------------------------------------")
            elif (first_match and "Logical" in config_mode):
                print ("!--------------------------------------------")
                print ("! Interface configuration (Logical) ")
                print ("!--------------------------------------------")
            first_match = False

            if self.is_portchannel_member(device_name,interface):
                pc_type = interface_list[device_name][interface]["PC-Type"]
                if "layer2" in pc_type:
                    print ("!....................................")
                    print ("!  Layer 2 PC: create physical first")
                    print ("!....................................")
                elif "layer3" in pc_type:
                    print ("!....................................")
                    print ("!  Layer 3 PC: create logical first")
                    print ("!....................................")
            # ---------------------------------------------------
            # Start generating interface specific configuration
            # ---------------------------------------------------
            print ("interface {}".format(interface_list[device_name][interface]["Interface"]))

            if self.is_portchannel_parent(device_name,interface):
                pc_members = interface_list[device_name][interface]["PC-Members"]
                print ("  !- pc members: {}".format(", ".join(pc_members)))
            if self.is_switch_port(device_name,interface):
                print ("  switchport")
            if self.is_routed_port(device_name,interface):
                print ("  no switchport")
            if self.has_description_configured(device_name,interface):
                print ("  description {}".format(interface_list[device_name][interface]["Description"]))
            if self.has_mtu_configured(device_name,interface):
                print ("  mtu {}".format(interface_list[device_name][interface]["MTU"]))
            if self.has_vrf_configured(device_name,interface):
                print ("  ip vrf forwarding {}".format(interface_list[device_name][interface]["VRF (leave blank if global)"]))
            if self.has_ip_configured(device_name,interface):
                print ("  ip address {}".format(self.GetIP(interface_list[device_name][interface]["IP Address (x.x.x.x/x)"])))
            if self.is_trunk_port(device_name,interface):
                trunk_vlans = interface_list[device_name][interface]["Trunk Allowed VLANs (separated by commas)"]
                print ("  switchport mode trunk")
                print ("  switchport trunk allowed vlan {}".format(trunk_vlans))
            if self.has_nativevlan_configured(device_name,interface):
                native_vlan = interface_list[device_name][interface]["Trunk Native VLAN"]
                print ("  switchport trunk native vlan {}".format(native_vlan))
            if self.is_data_port(device_name,interface):
                print ("  switchport access vlan {}".format(interface_list[device_name][interface]["Data VLAN"]))
            if self.is_voice_port(device_name,interface):
                print ("  switchport voice vlan {}".format(interface_list[device_name][interface]["Voice VLAN"]))
            if self.is_portchannel_member(device_name,interface):
                pc_group = interface_list[device_name][interface]["PC-Group"]
                pc_mode = interface_list[device_name][interface]["PC-Mode"]
                print ("  channel-group {} mode {}".format(pc_group,pc_mode))
            if self.has_variable1_configured(device_name,interface):
                if self.is_valid_variable(interface_list[device_name][interface]["Variable 1"]):
                    print (self.GetVariable(interface_list[device_name][interface]["Variable 1"]))
            if self.has_variable2_configured(device_name,interface):
                if self.is_valid_variable(interface_list[device_name][interface]["Variable 2"]):
                    print (self.GetVariable(interface_list[device_name][interface]["Variable 2"]))
            if self.has_speed_configured(device_name,interface):
                print ("  speed {}".format(interface_list[device_name][interface]["Speed"]))
            if self.has_duplex_configured(device_name,interface):
                print ("  duplex {}".format(interface_list[device_name][interface]["Duplex"]))
            if self.is_interface_enabled(device_name,interface):
                print ("  no shutdown")
            else:
                print ("  shutdown")
        else:
            print ("!")


    def GenerateConfig(self):
        global device_list
        console = sys.__stdout__
        print ("Generating configuration....")
        for device in device_list:
            sys.stdout = Logger(device+".txt")
            print ("****************************************")
            print ("! Device configuration for {}".format(device))
            print ("****************************************")
            self.CreateGlobalConfig(device,"Start")
            self.CreateVrfConfig(device)
            self.CreateVlanConfig(device)
            self.CreateInterfaceConfig(device,"Physical")
            self.CreateInterfaceConfig(device,"Logical")
            self.CreatePrefixConfig(device)
            self.CreateStaticRouteConfig(device)
            self.CreateGlobalConfig(device,"End")
            sys.stdout = console
            print ("- {} configuration generated.".format(device))

    def CheckInterfacesForErrors(self):
        for device in interface_list:
            for interface in sorted(interface_list[device]):
                if self.is_routed_port(device,interface) and self.is_switch_port(device,interface):
                    error_db["interfaces"].append("Row ({}): [{}] [{}] both routed and switchport config detected".format(interface_list[device][interface]["Row"],device,interface))
                if self.has_variable1_configured(device,interface):
                    if not self.GetVariable(interface_list[device][interface]["Variable 1"]):
                        error_db["interfaces"].append("Row ({}): [{}] [{}] referenced variable '{}' which does not exist".format(interface_list[device][interface]["Row"],device,interface,interface_list[device][interface]["Variable 1"]))
                if self.has_variable2_configured(device,interface):
                    if not self.GetVariable(interface_list[device][interface]["Variable 2"]):
                        error_db["interfaces"].append("Row ({}): [{}] [{}] referenced variable '{}' which does not exist".format(interface_list[device][interface]["Row"],device,interface,interface_list[device][interface]["Variable 2"]))
                if self.is_data_port(device,interface):
                    if not self.is_valid_vlan(device,interface_list[device][interface]["Data VLAN"]):
                        error_db["interfaces"].append("Row ({}): [{}] [{}] referenced Data VLAN '{}' which does not exist".format(interface_list[device][interface]["Row"],device,interface,interface_list[device][interface]["Data VLAN"]))
                if self.is_voice_port(device,interface):
                    if not self.is_valid_vlan(device,interface_list[device][interface]["Voice VLAN"]):
                        error_db["interfaces"].append("Row ({}): [{}] [{}] referenced Voice VLAN '{}' which does not exist".format(interface_list[device][interface]["Row"],device,interface,interface_list[device][interface]["Data VLAN"]))
                if self.has_nativevlan_configured(device,interface):
                    if not self.is_valid_vlan(device,interface_list[device][interface]["Trunk Native VLAN"]):
                        error_db["interfaces"].append("Row ({}): [{}] [{}] referenced Native VLAN '{}' which does not exist".format(interface_list[device][interface]["Row"],device,interface,interface_list[device][interface]["Trunk Native VLAN"]))
                if self.has_vrf_configured(device,interface):
                    if not self.is_valid_vrf(device,interface_list[device][interface]["VRF (leave blank if global)"]):
                        error_db["interfaces"].append("Row ({}): [{}] [{}] referenced vrf '{}' which does not exist".format(interface_list[device][interface]["Row"],device,interface,interface_list[device][interface]["VRF (leave blank if global)"]))
                if self.has_ip_configured(device,interface):
                    if not self.is_valid_ipaddress(interface_list[device][interface]["IP Address (x.x.x.x/x)"]):
                        error_db["interfaces"].append("Row ({}): [{}] [{}] using IPAddr '{}' which is invalid".format(interface_list[device][interface]["Row"],device,interface,interface_list[device][interface]["IP Address (x.x.x.x/x)"]))
                if self.is_trunk_port(device,interface):
                    if not self.is_valid_trunk(device,interface):
                        error_db["interfaces"].append("Row ({}): [{}] [{}] one or more vlans referenced in trunk do not exist".format(interface_list[device][interface]["Row"],device,interface))

    def GenerateErrorReport(self):
        console = sys.__stdout__
        sys.stdout = Logger("ccg-errors.txt")

        if error_db["profiles"]:
            print ("===========================")
            print ("Worksheet: [profiles]")
            print ("===========================")
            for entry in error_db["profiles"]:
                print (entry)

        if error_db["config-templates"]:
            print ("===========================")
            print ("Worksheet: [config-templates]")
            print ("===========================")
            for entry in error_db["config-templates"]:
                print (entry)

        if error_db["interfaces"]:
            print ("===========================")
            print ("Worksheet: [interfaces]")
            print ("===========================")
            for entry in error_db["interfaces"]:
                print (entry)
        sys.stdout = console

def StartCode():
    #--------------------------------------------------------------------------------------------------------
    # Execute the code
    #--------------------------------------------------------------------------------------------------------
    db = Config()               # Read the build spreadsheet and build the raw database
    db.RemoveEmptyRowsFromDB()  # Clean up the database and remove rows that don't have required columns
    #--------------------------------------------------------------------------------------------------------
    db.GetDeviceList()          # Scan through all the worksheets and capture a list of unique devices names
    db.GetVlanList()            # Scan through the vlan worksheet and capture a list of vlans per device
    db.GetVariableList()        # Scan through the variable worksheet and capture valid variables
    db.GetConfigTemplateList()  # Scan through the "config-templates" worksheet and capture templates
    db.GetProfileList()         # Scan through the "profiles" worksheet and capture which proilfes are used
    db.GetVrfList()             # Scan through the "vrf" worksheet and capture the VRFs
    db.GetInterfaceList()       # Scan through the "interfaces" worksheet and capture interfaces
    db.GetStaticRouteList()     # Scan through the "static routes" worksheet and capture routes
    db.GetPrefixList()          # Scan through the "prefix-list" worksheet and capture prefix-lists
    db.GetPortChannelList()     # Scan through the "portchannels" worksheet and capture all the portchannels
    #--------------------------------------------------------------------------------------------------------
    db.CheckInterfacesForErrors()
    #--------------------------------------------------------------------------------------------------------
    db.GenerateConfig()         # Generate the actual configuration
    db.GenerateErrorReport()
    #--------------------------------------------------------------------------------------------------------
    print ("\nConfiguration has been generated.")

#filename = "build-v2.0.xlsx"
#StartCode()

#---------------
# Show the menu
#---------------
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
        StartCode()
    except IOError:
        print ("Unable to open: %s"% sys.argv[1])
        print ("Program aborted.")
        exit()

if __name__ == '__main__':
    main(sys.argv)




# Shows the "global config" even if no entry has been found, fix this