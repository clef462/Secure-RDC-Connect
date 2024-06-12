import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import json
import pyautogui
import time
import threading

config_file = "config.json"


def load_last_selections():
    try:
        if os.path.exists(config_file):
            with open(config_file, 'r') as file:
                return json.load(file)
    except Exception as e:
        print(f"Error loading last selections: {e}")
    return {"last_vpn": "", "last_rdc": "", "rdc_directory": "C:\\Windows\\Temp\\RDC", "rasphone_path": ""}


def save_last_selections(vpn, rdc, rdc_directory, rasphone_path):
    try:
        with open(config_file, 'w') as file:
            json.dump(
                {"last_vpn": vpn, "last_rdc": rdc, "rdc_directory": rdc_directory, "rasphone_path": rasphone_path},
                file)
    except Exception as e:
        print(f"Error saving selections: {e}")


def fetch_vpn_connections(rasphone_path):
    vpn_list = []
    if rasphone_path:
        try:
            with open(rasphone_path, 'r') as file:
                for line in file:
                    if line.startswith('ENTRYNAME'):
                        vpn_list.append(line.split('=')[1].strip())
        except Exception as e:
            print(f"Error fetching VPN connections: {e}")
    return vpn_list


def fetch_rdc_connections(directory):
    print(f"Fetching RDC connections from {directory}")
    rdc_list = []
    try:
        for file in os.listdir(directory):
            if file.endswith(".rdp"):
                rdc_list.append(file)
    except Exception as e:
        print(f"Error fetching RDC connections: {e}")
    print(f"Found RDC connections: {rdc_list}")
    return rdc_list


def select_rasphone_file():
    try:
        rasphone_path = filedialog.askopenfilename(filetypes=[("Phonebook Files", "*.pbk")])
        if rasphone_path:
            print(f"Selected rasphone.pbk file: {rasphone_path}")
            rasphone_path_var.set(rasphone_path)
            update_vpn_connections()
    except Exception as e:
        print(f"Error selecting rasphone.pbk file: {e}")


def select_rdc_directory():
    try:
        directory = filedialog.askdirectory()
        if directory:
            print(f"Selected RDC directory: {directory}")
            rdc_var.set("")
            rdc_directory_var.set(directory)
            update_rdc_connections()
    except Exception as e:
        print(f"Error selecting RDC directory: {e}")


def update_vpn_connections():
    try:
        rasphone_path = rasphone_path_var.get()
        if rasphone_path:
            vpn_connections = fetch_vpn_connections(rasphone_path)
            set_vpn_connections(vpn_connections)
    except Exception as e:
        print(f"Error in update_vpn_connections: {e}")


def set_vpn_connections(vpn_connections):
    try:
        vpn_dropdown['values'] = vpn_connections
        print(f"VPN connections updated: {vpn_connections}")
    except Exception as e:
        print(f"Error setting VPN connections: {e}")


def update_rdc_connections():
    try:
        rdc_directory = rdc_directory_var.get()
        if rdc_directory:
            rdc_connections = fetch_rdc_connections(rdc_directory)
            set_rdc_connections(rdc_connections)
    except Exception as e:
        print(f"Error in update_rdc_connections: {e}")


def set_rdc_connections(rdc_connections):
    try:
        rdc_dropdown['values'] = rdc_connections
        print(f"RDC connections updated: {rdc_connections}")
    except Exception as e:
        print(f"Error setting RDC connections: {e}")


def connect_vpn_and_rdc():
    def handle_connect():
        try:
            selected_vpn = vpn_var.get()
            rasphone_path = rasphone_path_var.get()

            if selected_vpn and rasphone_path:
                print(f"Connecting to VPN: {selected_vpn}")
                subprocess.Popen(['rasphone', '-d', selected_vpn])

                time.sleep(3)  # Wait for the rasphone window to appear

                try:
                    # Automate the pressing of the ENTER key for both connect buttons
                    pyautogui.press('enter')
                    time.sleep(1)
                    pyautogui.press('enter')
                    time.sleep(10)  # Wait for VPN connection to establish
                except Exception as e:
                    print(f"Error sending ENTER key: {e}")

                print(f"Connected to VPN: {selected_vpn}")

                selected_rdc = rdc_var.get()
                rdc_directory = rdc_directory_var.get()
                rdc_path = os.path.join(rdc_directory, selected_rdc)

                if selected_rdc:
                    print(f"Connecting to RDC: {selected_rdc}")
                    subprocess.Popen(['mstsc', rdc_path])
                    print(f"Connected to RDC: {selected_rdc}")

                save_last_selections(selected_vpn, selected_rdc, rdc_directory, rasphone_path)
                show_disconnect_button()
        except Exception as e:
            print(f"Error in connect_vpn_and_rdc: {e}")
            messagebox.showerror("Connection Error", f"Failed to connect: {e}")

    threading.Thread(target=handle_connect).start()


def show_disconnect_button():
    try:
        connect_button.pack_forget()
        disconnect_button.pack(pady=10)
        disconnect_button.lift()
    except Exception as e:
        print(f"Error showing disconnect button: {e}")


def disconnect_vpn_and_rdc():
    try:
        selected_vpn = vpn_var.get()
        if selected_vpn:
            print(f"Disconnecting from VPN: {selected_vpn}")
            subprocess.run(['rasdial', selected_vpn, '/disconnect'], check=True)
            print(f"Disconnected from VPN: {selected_vpn}")
    except Exception as e:
        print(f"Error in disconnect_vpn_and_rdc: {e}")
        messagebox.showerror("Disconnection Error", f"Failed to disconnect: {e}")
    disconnect_button.pack_forget()
    connect_button.pack(pady=10)


# Load last selections
last_selections = load_last_selections()

# Initialize Tkinter
root = tk.Tk()
root.title("Secure-RDC-Connect")
root.geometry("440x360")  # Set the initial size to 440 pixels wide by 360 pixels tall

# Title label
title_label = tk.Label(root, text="Secure-RDC-Connect", font=("Helvetica", 16, "bold"))
title_label.pack(pady=10)

# Fetch available VPN connections
rasphone_path_var = tk.StringVar(value=last_selections.get("rasphone_path", ""))
vpn_connections = fetch_vpn_connections(rasphone_path_var.get())

# Create drop-down for VPN
vpn_var = tk.StringVar(value=last_selections["last_vpn"])
vpn_label = tk.Label(root, text="Select VPN:")
vpn_label.pack(pady=5)
vpn_dropdown = ttk.Combobox(root, textvariable=vpn_var, values=vpn_connections)
vpn_dropdown.pack(pady=5)

# Create drop-down for RDC
rdc_var = tk.StringVar(value=last_selections["last_rdc"])
rdc_directory_var = tk.StringVar(value=last_selections.get("rdc_directory", "C:\\Windows\\Temp\\RDC"))
rdc_label = tk.Label(root, text="Select RDC:")
rdc_label.pack(pady=5)
rdc_dropdown = ttk.Combobox(root, textvariable=rdc_var)
rdc_dropdown.pack(pady=5)

# Button to select the rasphone.pbk file
select_rasphone_button = tk.Button(root, text="Select rasphone.pbk", command=select_rasphone_file)
select_rasphone_button.pack(pady=5)

# Button to select the directory containing RDP files
select_directory_button = tk.Button(root, text="Select RDC Directory", command=select_rdc_directory)
select_directory_button.pack(pady=5)

# Create button to connect
connect_button = tk.Button(root, text="Connect", command=connect_vpn_and_rdc, font=("Helvetica", 12), fg="green",
                           width=20, height=2)
connect_button.pack(pady=10)

# Create a disconnect button
disconnect_button = tk.Button(root, text="Disconnect", command=disconnect_vpn_and_rdc, font=("Helvetica", 12), bg="red",
                              fg="white", width=20, height=2)
disconnect_button.pack_forget()

# Footer separator
separator = ttk.Separator(root, orient='horizontal')
separator.pack(fill='x', pady=5)

# Footer label
footer_label = tk.Label(root, text="Copyright by K.Bussard", font=("Helvetica", 10))
footer_label.pack(pady=5)

# Ensure the disconnect button stays on top
root.attributes('-topmost', True)

# Update VPN and RDC connections based on the last selected paths
update_vpn_connections()
update_rdc_connections()

# Run the Tkinter event loop
root.mainloop()
