import subprocess
import json
import socket
import re
import ipaddress
import psutil

def get_device_name(ip):
    try:
        return socket.gethostbyaddr(ip)[0]
    except socket.herror:
        return "Unknown"

def get_local_networks():
    networks = []
    for interface, snics in psutil.net_if_addrs().items():
        for snic in snics:
            if snic.family == socket.AF_INET:
                ip = ipaddress.ip_address(snic.address)
                if not ip.is_loopback:
                    # Assuming a /24 subnet mask, which is common.
                    # This might need adjustment for different network configurations.
                    network = ipaddress.ip_network(f'{snic.address}/24', strict=False)
                    networks.append(network)
    return networks

def scan_network():
    try:
        # Execute the 'arp -a' command
        arp_output = subprocess.check_output("arp -a", shell=True).decode('utf-8')
    except subprocess.CalledProcessError:
        return []

    devices = []
    local_networks = get_local_networks()
    # Regex to find IP and MAC addresses
    pattern = re.compile(r"(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s+([0-9a-fA-F]{2}(?:[:-][0-9a-fA-F]{2}){5})")
    
    for line in arp_output.split('\n'):
        match = pattern.search(line)
        if match:
            ip_address_str = match.group(1)
            ip_address = ipaddress.ip_address(ip_address_str)
            for network in local_networks:
                if ip_address in network:
                    mac_address = match.group(2)
                    device_name = get_device_name(ip_address_str)
                    devices.append({
                        "ip": ip_address_str,
                        "mac": mac_address,
                        "name": device_name
                    })
                    break # Move to the next line once a match is found
    return devices

if __name__ == "__main__":
    devices = scan_network()
    with open("network_devices.json", "w") as f:
        json.dump(devices, f, indent=4)
