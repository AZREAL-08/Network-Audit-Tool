## *Network Inventory Script*

### *Project Overview*

This Python-based script automates the process of gathering inventory data (hardware specs, installed software, configuration details) from remote Windows computers on your network.
It uses *Windows Remote Management (WinRM)* and *PowerShell Remoting* to connect to systems and execute commands.
The collected data is then compiled and exported into an *Excel file* for easy management, tracking, and analysis.
This tool is designed for IT professionals and system administrators.

---

### *Features*

* Retrieves hardware and software information from remote computers.
* Supports exporting all collected data to a single Excel file.
* Designed for Windows-based machines using PowerShell Remoting (PSRemoting).

---

### *Requirements*

You will need the following Python packages. The os package is built-in and does not require installation.

* *pypsrp:* To execute PowerShell commands over WinRM.
* *pandas:* For handling and structuring the collected data.
* *openpyxl:* For writing data to Excel files.

You can install all necessary packages using pip:

bash
pip install pypsrp pandas openpyxl


---

### *Setup Instructions*

Before running the script, you must configure the target remote machines and the machine running the script.

#### *1. Configure Target Windows Machines*

The following steps must be performed on each remote machine you want to inventory. You will need administrator rights.

*A. Enable PowerShell Remoting*

Run the following command in PowerShell (as Administrator) to enable PSRemoting:

powershell
Enable-PSRemoting -Force


This allows the system to accept remote commands via WinRM.

*B. Configure TrustedHosts*

To allow your script's machine to connect, you must configure the TrustedHosts list on the target machine.

powershell
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force


⚠ *Security Warning:*
The command above (-Value "*") allows connections from any machine.
For a more secure environment, replace * with the specific IP address of the machine that will be running the inventory script.

*C. Configure Firewall*

Ensure the Windows Firewall allows WinRM traffic.
By default, this is port *5985 (HTTP)* or *5986 (HTTPS)*.

You can use the following PowerShell command to create a firewall rule for the default HTTP port:

powershell
New-NetFirewallRule -DisplayName "Allow WinRM (HTTP-In)" -Name "WinRM-HTTP-In" -Protocol TCP -LocalPort 5985 -Action Allow


---

#### *2. Network Requirements*

Ensure that the machine running the script and all target machines are connected to the same subnet.
Communication across different subnets may fail without additional network configurations (like VPNs or specific routing rules).

---

#### *3. Create Credentials File (.env)*

To securely store the admin credentials for the remote systems, create a file named **.env** in the same directory as the script.
The script will read this file to authenticate with each target machine.

**Example .env File Format:**


TARGET_1_USERNAME=admin_user
TARGET_1_PASSWORD=AdminPassword123
TARGET_2_USERNAME=admin_user2
TARGET_2_PASSWORD=AdminPassword456


*Important:*
The credentials provided in the .env file must have administrator rights on the corresponding target machines to successfully collect all inventory data.

---

### *Running the Script*

Once all configurations are in place, you can run the script from your terminal:

bash
python network_inventory.py


This will initiate connections, collect the inventory data, and save the results in a new Excel file in the project directory.

---

### *Warnings*

* *Network Configuration:* All machines must be reachable within the same subnet.
* *Firewall Settings:* Incorrect firewall settings are the most common cause of connection failure. Double-check that ports *5985/5986* are open on the target machines.
* *Admin Credentials:* The script will fail to gather data if the credentials in the .env file do not have administrator privileges on the remote systems.
