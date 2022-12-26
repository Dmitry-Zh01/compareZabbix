Hello, everyone. This script allows to automate comparing hosts from two zabbix servers and creation Excel report.

<h2><strong>Information:</strong></h2>
<p>Complete Excel report made using this script should consist of two worksheets: information about all hosts from the first Zabbix Server and information about all hosts from the second Zabbix Server.<p>

![image](https://user-images.githubusercontent.com/106164393/209563232-42d80409-4582-4991-952f-70ad4382a136.png)

<h2><strong>The following are required for proceeding this script:</strong></h2>

<strong>python3</strong>
<strong>modules:</strong> sys os json pyzabbix getpass3 openpyxl platform subprocess time datetime collections progress

<strong>How to?</strong> To executing the script and get report you should run this and follow promts from python terminal:

Add full 1 Zabbix API path, such as: https://zabbixaddress/api_jsonrpc.php or https://zabbixaddress/zabbix/api_jsonrpc.php<br>
Type Zabbix API username and Zabbix API user password<br>
Then, set 2 Zabbix API path, such as: https://zabbixaddress2/api_jsonrpc.php or https://zabbixaddress2/zabbix/api_jsonrpc.php<br>
Add 2 Zabbix API username and 2 Zabbix API user password<br>
Finally, choose timeframes (from and till datetime in the format: 01/01/2000 00:00) for getting SLA info via hosts ICMP ping history or press 'Enter' to set default values (30 days ago from the moment of script running)<br>

Output data:<br>
Folder: <strong>Excel</strong><br>
    - <strong>Zabbix_compare.xlsx</strong> <-- Excel report with two sheets<br>
    - <strong>new.json</strong>            <-- JSON file with information from the first Zabbix server<br>
    - <strong>old.json</strong>            <-- JSON file with information from the second Zabbix server<br>

Good luck! :)
