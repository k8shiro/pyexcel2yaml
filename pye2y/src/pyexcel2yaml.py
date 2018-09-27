# -*- coding: utf-8 -*-
import codecs, sys
import xlrd
import yaml
import re
import os

class Excel_2_Yaml():
    def __init__(self, excel_path, export_path):
        self.set_book(excel_path)
        self.set_sheets()
        self.export_path = export_path
        self.hostvars = {
            'ansible': {},
            'serverspec':{},
        }
        self.inventory = {
            'ansible': [], 
            'serverspec': []
        }

    def set_book(self, excel_path):
        self.book = xlrd.open_workbook(excel_path)


    def set_sheets(self):
        self.sheets = {self.book.sheet_by_index(num).name: self.book.sheet_by_index(num) for num in range(self.book.nsheets)}


    def parse_inventory(self):
        inventory_sheet = self.sheets['Inventory']
        for row in range(19, inventory_sheet.nrows):
            host_setting = {
                'name': inventory_sheet.cell(row, 2).value,
                'conn_user': inventory_sheet.cell(row, 3).value,
                'conn_password': inventory_sheet.cell(row, 4).value,
                'roles': [],
            }

            for col in range(7, inventory_sheet.ncols):
                if (not inventory_sheet.cell(18, col).value =='Ansible'):
                    continue

                if (inventory_sheet.cell(row, col).value =='○'):
                    host_setting['roles'].append(inventory_sheet.cell(17, col).value)

            if (not host_setting['roles'] == []):
                self.inventory['ansible'].append(host_setting)

        for row in range(19, inventory_sheet.nrows):
            host_setting = {
                'name': inventory_sheet.cell(row, 2).value,
                'conn_user': inventory_sheet.cell(row, 3).value,
                'conn_password': inventory_sheet.cell(row, 4).value,
                'roles': [],
            }

            for col in range(7, inventory_sheet.ncols):
                if (not inventory_sheet.cell(18, col).value =='Serverspec'):
                    continue

                if (inventory_sheet.cell(row, col).value =='○'):
                    host_setting['roles'].append(inventory_sheet.cell(17, col).value)

            if (not host_setting['roles'] == []):
                self.inventory['serverspec'].append(host_setting)



    def export_serverspec_inventory(self):
        serverspec_inventory = self.inventory['serverspec']

        f = open(os.path.join(self.export_path, 'Serverspec.1.inventory'), 'w')
        f.write(yaml.safe_dump(serverspec_inventory, default_flow_style=False, allow_unicode=True))
        f.close()

    def export_ansible_inventory(self):
        ansible_inventory = {}
        for target in self.inventory['ansible']:
            for role in target['roles']:
                if not role in ansible_inventory:
                    ansible_inventory[role] = []
                ansible_inventory[role].append({
                    'target_name': target['name'],
                    'ansible_user': target['conn_user'],
                    'ansible_ssh_pass': target['conn_password']
                })

        ansible_inventory_to_text = ''

        for role in ansible_inventory:
            add_text = "[ {}_GROUP ]\n".format(role)
            ansible_inventory_to_text += add_text

            for target in ansible_inventory[role]:
                add_text = "{} ansible_user={} ansible_ssh_pass={}\n".format(
                                                                target['target_name'],
                                                                target['ansible_user'],
                                                                target['ansible_ssh_pass']
                )
                ansible_inventory_to_text += add_text


        add_text = "[linux:children]\n"
        ansible_inventory_to_text += add_text

        for role in ansible_inventory:
            if re.match("1-" , role):
                add_text = "{}_GROUP\n".format(role)
                ansible_inventory_to_text += add_text

        add_text = "[windows:children]\n"
        ansible_inventory_to_text += add_text

        for role in ansible_inventory:
            if re.match("2-" , role):
                add_text = "{}_GROUP\n".format(role)
                ansible_inventory_to_text += add_text

        other_settings = (
            '[linux:vars] \n'
            'ansible_ssh_port = 22 \n'
            '[windows:vars] \n'
            'ansible_ssh_port = 5986 \n'
            'ansible_connection = winrm \n'
            'ansible_winrm_server_cert_validation = ignore \n'
        )

        ansible_inventory_to_text += other_settings

        f = open(os.path.join(self.export_path, 'Ansible.1.inventory'), 'w')
        f.write(ansible_inventory_to_text)
        f.close() 


    def parse_parameter_sheets(self):
        parameter_sheets = []
        for seet_name in self.sheets:
            if seet_name != 'Inventory':
                parameter_sheets.append(self.sheets[seet_name])

        for parameter_sheet in parameter_sheets:
            # anssible
            yaml_data = ''
            for row in range(18, parameter_sheet.nrows):
                key = parameter_sheet.cell(row, 13).value
                key_type = parameter_sheet.cell(row, 14).value
                value = parameter_sheet.cell(row, 9).value
                
                if key_type != 'Object' and value == '':
                    continue
                if value != '':
                    value = "'{}'".format(value)

                yaml_col = "{}: {}".format(key, value)

                yaml_data += yaml_col + '\n'
            self.hostvars['ansible'][parameter_sheet.name] = yaml.load(yaml_data)

            # serverspec
            yaml_data = ''
            for row in range(18, parameter_sheet.nrows):
                key = parameter_sheet.cell(row, 13).value
                value = parameter_sheet.cell(row, 10).value
                if value != '':
                    value = "'{}'".format(value)
                yaml_col = "{}: {}".format(key, value)

                yaml_data += yaml_col + '\n'
            self.hostvars['serverspec'][parameter_sheet.name] = yaml.load(yaml_data)


    def del_null_vars(self, target_dict):
        filtered_dict = target_dict
        for key, val in target_dict.items():
            if val == None:
                filtered_dict.pop(key)
            elif isinstance(val, dict):
                filtered_dict[key] = self.del_null_vars(val)
        return filtered_dict

    def export_ansible_hostvars(self):
        for sheetname, vars in self.hostvars['ansible'].items():
            vars = self.del_null_vars(vars)
            connection_hostname = vars['connection_hostname']
            f = open(os.path.join(self.export_path, '{}.yml'.format(connection_hostname)), 'w')
            f.write(yaml.dump(vars, default_flow_style=False, allow_unicode=True))
            f.close()

    def export_serverspec_hostvars(self):
        properties = {}
        for sheetname, vars in self.hostvars['ansible'].items():
            vars = self.del_null_vars(vars)
            connection_hostname = vars['connection_hostname']
            properties[connection_hostname] = vars


        f = open(os.path.join(self.export_path, 'properties.yml'), 'w')
        f.write(yaml.dump(properties, default_flow_style=False, allow_unicode=True))
        f.close()


def main():
    reload(sys)
    sys.setdefaultencoding('utf-8')
    sys.stdout = codecs.getwriter('utf_8')(sys.stdout)
    sys.stdin = codecs.getreader('utf_8')(sys.stdin)

    excel2yaml = Excel_2_Yaml('./Excel2YAML_1.0.0.xlsm', './export_files')

    excel2yaml.parse_inventory()
    print("============INVENTORY:ANSIBLE=========")
    excel2yaml.export_ansible_inventory()
    print("============INVENTORY:SERVERSPEC======")
    excel2yaml.export_serverspec_inventory()

    excel2yaml.parse_parameter_sheets()
    print("============HTOSVARS:ANSIBLE=========")
    excel2yaml.export_ansible_hostvars()
    print("============HTOSVARS:SERVERSPEC=========")
    excel2yaml.export_serverspec_hostvars()

if __name__ == '__main__':
    main()





