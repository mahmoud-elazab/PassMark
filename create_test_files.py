"""
Interactive test file generator for PassMark Tool.
Creates dummy .txt files for all selected modes and devices.
Run: python create_test_files.py
"""

import os
import random


OUTPUT_FOLDER = "test_files"

# Voltage ranges per mode per port (min, max) for realistic dummy data
VOLTAGE_RANGES = {
    2: {1: (14.90, 15.10), 2: (4.90, 5.10)},
    3: {1: (4.90, 5.10),  2: (14.90, 15.10)},
    4: {1: (4.90, 5.10),  2: (4.90, 5.10)},
    5: {1: (14.90, 15.10), 2: (14.90, 15.10)},
    6: {1: None,           2: (19.90, 20.10)},
    7: {1: (19.90, 20.10), 2: None},
}

# Which ports are active per mode
ACTIVE_PORTS = {
    2: [1, 2],
    3: [1, 2],
    4: [1, 2],
    5: [1, 2],
    6: [2],
    7: [1],
}


def make_file(filename, voltage_range, num_readings=10):
    """
    Create a PassMark log file with random voltage readings
    within the given range.
    """
    lines = [
        "PASSMARK USB Power Delivery Tester Log file",
        "===========================================",
        f"Number of records logged: {num_readings + 1}",
        "",
        "Tue Apr  7 11:41:41 2026: Information- "
        "PassMark Power Delivery Tester V1.1 1024",
    ]

    for i in range(num_readings):
        voltage = float(f"{random.uniform(voltage_range[0], voltage_range[1]):.2f}")
        lines.append(
            f"Tue Apr  7 11:41:{42 + i} 2026: "
            f"Diagnostics- V={voltage:.2f}V, I Tester=3000mA, "
            f"I Loopback=0mA, I Total~=3000mA"
        )

    content = "\n".join(lines)
    path = os.path.join(OUTPUT_FOLDER, filename)
    with open(path, "w") as f:
        f.write(content)

    print(f"  Created: {filename}")


def get_devices():
    """Ask user how many devices and their numbers."""
    print("\n" + "=" * 50)
    print("  PassMark Tool — Test File Generator")
    print("=" * 50)

    while True:
        try:
            count = int(input("\nHow many devices do you want to create files for? "))
            if count < 1:
                print("Please enter at least 1.")
                continue
            break
        except ValueError:
            print("Please enter a valid number.")

    devices = []
    print()
    for i in range(count):
        while True:
            device = input(f"  Enter device number {i + 1}: ").strip()
            if device == "":
                print("  Device number cannot be empty.")
                continue
            if device in devices:
                print(f"  Device {device} already entered. Please use a different number.")
                continue
            devices.append(device)
            break

    return devices


def get_modes():
    """Ask user which modes to generate files for."""
    print("\nAvailable Operating Modes:")
    print("  2 → 45W/15W  (Port 1 = 15V, Port 2 = 5V)")
    print("  3 → 15W/45W  (Port 1 = 5V,  Port 2 = 15V)")
    print("  4 → 15W/15W  (Port 1 = 5V,  Port 2 = 5V)")
    print("  5 → 30W/30W  (Port 1 = 15V, Port 2 = 15V)")
    print("  6 → 0W/60W   (Port 2 = 20V only)")
    print("  7 → 60W/0W   (Port 1 = 20V only)")

    print("\nEnter the modes you want (e.g: 2 3 4 5 6 7)")
    print("Or press Enter to select all modes (2,3,4,5,6,7):")

    while True:
        raw = input("  Modes: ").strip()

        if raw == "":
            selected = [2, 3, 4, 5, 6, 7]
            break

        try:
            selected = [int(m) for m in raw.split()]
            invalid = [m for m in selected if m not in VOLTAGE_RANGES]
            if invalid:
                print(f"  Invalid modes: {invalid}. Please choose from 2 to 7.")
                continue
            if len(selected) == 0:
                print("  Please select at least one mode.")
                continue
            break
        except ValueError:
            print("  Please enter valid mode numbers separated by spaces.")

    return sorted(set(selected))


def get_num_readings():
    """Ask user how many readings per file."""
    print("\nHow many voltage readings per file?")
    print("Press Enter for default (10):")

    while True:
        raw = input("  Readings: ").strip()
        if raw == "":
            return 10
        try:
            n = int(raw)
            if n < 1:
                print("  Please enter at least 1.")
                continue
            return n
        except ValueError:
            print("  Please enter a valid number.")


def confirm(devices, modes, num_readings):
    """Show summary and ask for confirmation."""
    total_files = 0
    for mode in modes:
        ports = ACTIVE_PORTS[mode]
        total_files += len(devices) * len(ports) * 2  # 2 angles per port

    print("\n" + "=" * 50)
    print("  Summary")
    print("=" * 50)
    print(f"  Devices:       {', '.join(devices)}")
    print(f"  Modes:         {', '.join(str(m) for m in modes)}")
    print(f"  Readings/file: {num_readings}")
    print(f"  Total files:   {total_files}")
    print(f"  Output folder: {OUTPUT_FOLDER}/")
    print("=" * 50)

    while True:
        answer = input("\nProceed? (y/n): ").strip().lower()
        if answer == 'y':
            return True
        if answer == 'n':
            return False
        print("Please enter y or n.")


def main():
    devices = get_devices()
    modes = get_modes()
    num_readings = get_num_readings()

    if not confirm(devices, modes, num_readings):
        print("\nCancelled. No files were created.")
        return

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    print(f"\nCreating files in '{OUTPUT_FOLDER}/'...\n")

    total_created = 0

    for device in devices:
        print(f"Device {device}:")

        for mode in modes:
            ports = ACTIVE_PORTS[mode]

            for port in ports:
                voltage_range = VOLTAGE_RANGES[mode][port]

                if voltage_range is None:
                    continue

                for angle in [1, 2]:
                    filename = f"{device}-{mode}-{port}-{angle}.txt"
                    make_file(filename, voltage_range, num_readings)
                    total_created += 1

        print()

    print("=" * 50)
    print(f"  Done! {total_created} files created in '{OUTPUT_FOLDER}/'")
    print("=" * 50)


if __name__ == "__main__":
    main()