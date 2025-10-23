import pandas as pd
from pypsrp.client import Client
import socket
import ipaddress
from dotenv import load_dotenv
import os

# ====== Load credentials securely from .env ======
load_dotenv()
USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")

# ====== Subnet to scan (change as per your network) ======
SUBNET = "192.168.1.0/24"

# ====== Data storage lists ======
software_data = []
device_data = []
device_logs = []


def scan_network(subnet):
    """Scan subnet manually for devices with WinRM (5985) open."""
    alive_hosts = []
    for ip in ipaddress.IPv4Network(subnet, strict=False):
        ip = str(ip)
        try:
            # Try connecting to WinRM port (5985)
            with socket.create_connection((ip, 5985), timeout=0.5):
                alive_hosts.append(ip)
        except:
            pass
    return alive_hosts


def connect_and_fetch(ip):
    """Connect to remote Windows machine and collect data."""
    try:
        client = Client(ip, username=USERNAME, password=PASSWORD, ssl=False)

        # ---- Get username of remote user ----
        user, _, _ = client.execute_cmd("whoami")

        # ---- Get installed applications (from Registry) ----
        apps_script = r'''
        Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
        HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
        Select-Object -ExpandProperty DisplayName | Where-Object { $_ -ne $null }
        '''
        apps, _, _ = client.execute_ps(apps_script)

        for app in apps.strip().split("\n"):
            software_data.append({
                "IP": ip,
                "App Name": app.strip(),
                "User": user.strip()
            })

        # ---- Get currently connected USB devices ----
        usb_script = 'Get-PnpDevice -Class USB | Select-Object -ExpandProperty FriendlyName'
        devices, _, _ = client.execute_ps(usb_script)

        for dev in devices.strip().split("\n"):
            device_data.append({
                "IP": ip,
                "Device Type": dev.strip(),
                "User": user.strip()
            })

        # ---- Get historical USB connection logs ----
        logs_script = r'''
        Get-WinEvent -LogName Microsoft-Windows-DriverFrameworks-UserMode/Operational |
        Where-Object { $_.Id -in 2003,2100 } |
        Select-Object TimeCreated, Id, Message -First 50
        '''
        logs, _, _ = client.execute_ps(logs_script)

        for log in logs.strip().split("\n"):
            if log.strip():
                device_logs.append({
                    "IP": ip,
                    "Event": log.strip(),
                    "User": user.strip()
                })

    except Exception as e:
        print(f"❌ Failed to connect to {ip}: {e}")


def write_excel():
    """Write all collected data into an Excel file with 3 sheets."""
    df_software = pd.DataFrame(software_data, columns=["IP", "App Name", "User"])
    df_devices = pd.DataFrame(device_data, columns=["IP", "Device Type", "User"])
    df_logs = pd.DataFrame(device_logs, columns=["IP", "Event", "User"])

    with pd.ExcelWriter("network_inventory.xlsx", engine="openpyxl") as writer:
        df_software.to_excel(writer, sheet_name="Software Inventory", index=False)
        df_devices.to_excel(writer, sheet_name="Current USB Devices", index=False)
        df_logs.to_excel(writer, sheet_name="USB Device Logs", index=False)

    print("✅ Excel created: network_inventory.xlsx with 3 sheets")


if __name__ == "__main__":
    # ====== Main Execution Flow ======
    ips = scan_network(SUBNET)
    print(f"📡 Discovered Devices: {ips}")

    for ip in ips:
        print(f"🔗 Connecting to {ip}...")
        connect_and_fetch(ip)

    write_excel()
