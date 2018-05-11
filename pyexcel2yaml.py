# coding=utf-8
import codecs, sys
import xlrd
import yaml
import re


class Excel_2_Yaml():
    def __init__(self, excel_path):
        self.set_book(excel_path)
        self.set_sheets()


    def set_book(self, excel_path):
        self.book = xlrd.open_workbook(excel_path)


    def set_sheets(self):
        self.sheets = {self.book.sheet_by_index(num).name: self.book.sheet_by_index(num) for num in range(self.book.nsheets)}


    def parse_inventory(self):
        inventory_sheet = self.sheets['Inventory']
        self.inventory = {'ansible': [], 'serverspec': [] }
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

        #print(self.inventory)
        #print(yaml.safe_dump(self.inventory, default_flow_style=False ))


    def export_serverspec_inventory(self):
        serverspec_inventory = self.inventory['serverspec']
        print(yaml.safe_dump(serverspec_inventory, default_flow_style=False ))

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

        for role in ansible_inventory:
            print("[ {}_GROUP ]".format(role))

            for target in ansible_inventory[role]:
                print("{} ansible_user={} ansible_ssh_pass={}".format(
                                                                target['target_name'],
                                                                target['ansible_user'],
                                                                target['ansible_ssh_pass']
                ))

        print("[linux:children]")
        for role in ansible_inventory:
            if re.match("1-" , role):
                print("{}_GROUP".format(role))

        print("[windows:children]")
        for role in ansible_inventory:
            if re.match("2-" , role):
                print("{}_GROUP".format(role))

        other_settings = (
            '[linux:vars] \n'
            'ansible_ssh_port = 22 \n'
            '[windows:vars] \n'
            'ansible_ssh_port = 5986 \n'
            'ansible_connection = winrm \n'
            'ansible_winrm_server_cert_validation = ignore \n'
        )

        print(other_settings)


    def parse_parameter_sheets(self):
        parameter_sheets = []
        for seet_name in self.sheets:
            if seet_name != 'Inventory':
                parameter_sheets.append(self.sheets[seet_name])

        for parameter_sheet in parameter_sheets:
            yaml_data = ''
            for row in range(20, parameter_sheet.nrows):
                key = parameter_sheet.cell(row, 13).value
                value = parameter_sheet.cell(row, 9).value
                if value != '':
                    value = "'{}'".format(value)
                yaml_col = "{}: {}".format(key, value)

                yaml_data += yaml_col + '\n'
            print("============{}======================".format(parameter_sheet.name))
            print(yaml.load(yaml_data))


def main():
    reload(sys)
    sys.setdefaultencoding('utf-8')
    sys.stdout = codecs.getwriter('utf_8')(sys.stdout)
    sys.stdin = codecs.getreader('utf_8')(sys.stdin)

    excel2yaml = Excel_2_Yaml('./Excel2YAML_1.0.0.xlsm')

    excel2yaml.parse_inventory()
    print("============ANSIBLE=========")
    excel2yaml.export_ansible_inventory()
    print("============SERVERSPEC======")
    excel2yaml.export_serverspec_inventory()

    excel2yaml.parse_parameter_sheets()

if __name__ == '__main__':
    main()





