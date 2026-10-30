"""
Parses PassMark log files.
Extracts voltage readings from diagnostic lines.
"""

import re


def parse_voltage_readings(filepath):
    """
    Read a PassMark log file and extract all voltage readings.

    Args:
        filepath: Path to the .txt log file

    Returns:
        List of float values (voltages in Volts)
        Empty list if no readings found.
    """
    voltages = []

    # This pattern matches V=14.98V and captures the number part
    voltage_pattern = re.compile(r'V=(\d+\.\d+)V')

    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            # Skip any line that isn't a Diagnostics reading
            if 'Diagnostics-' not in line:
                continue

            match = voltage_pattern.search(line)
            if match:
                voltage = float(match.group(1))
                voltages.append(voltage)

    return voltages


def get_voltage_average(filepath):
    """
    Calculate the average voltage from a PassMark log file.

    Args:
        filepath: Path to the .txt log file

    Returns:
        Float: average voltage, or None if no readings found
    """
    voltages = parse_voltage_readings(filepath)

    if not voltages:
        print(f"Warning: No voltage readings found in {filepath}")
        return None

    average = sum(voltages) / len(voltages)
    return average