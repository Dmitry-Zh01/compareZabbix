# compareZabbix
This script allows to automate comparing hosts from two zabbix servers and creation Excel report.

Hello, everyone. This script allows to automate comparing hosts from two zabbix servers and creation Excel report.
Complete Excel report made using this script should consist of two worksheets: information about all hosts from the first Zabbix Server and information about all hosts from the second Zabbix Server.

![image](https://user-images.githubusercontent.com/106164393/209563232-42d80409-4582-4991-952f-70ad4382a136.png)

The following are required for proceeding this script:

python3
modules: sys os json pyzabbix getpass openpyxl platform subprocess time datetime collections progress

How to? To executing the script and get report you should run this and follow promts from python terminal:

Add full 1 Zabbix API path, such as: https://zabbixaddress/api_jsonrpc.php or https://zabbixaddress/zabbix/api_jsonrpc.php
Type 1 Zabbix API username
Type 1 Zabbix API user password
Then, set 2 Zabbix API path, such as: https://zabbixaddress2/api_jsonrpc.php or https://zabbixaddress2/zabbix/api_jsonrpc.php
Type 2 Zabbix API username
Type 2 Zabbix API user password
Choose timeframes (from and till datetime in the format: 01/01/2000 00:00) for getting SLA info via hosts ICMP ping history or press 'Enter' to set default values (30 days ago from the moment of script running)

Good luck! :)
