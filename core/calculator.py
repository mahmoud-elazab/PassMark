"""
Handles the pairing logic and combined average calculations.

Pairing rule:
- Files are grouped by (device, mode, port)
- If both angles exist → combined average
- If only one angle exists → individual average + warning
"""

from core.filename_parser import parse_filename
from core.file_parser import get_voltage_average


# Full definition of all supported operating modes
# port1/port2: None means unused (black cell in Excel)
MODE_DEFINITIONS = {
    2: {
        'label': '45W/15W',
        'port1': {'watts': 45, 'voltage': 15.0, 'current': 3},
        'port2': {'watts': 15, 'voltage': 5.0,  'current': 3},
    },
    3: {
        'label': '15W/45W',
        'port1': {'watts': 15, 'voltage': 5.0,  'current': 3},
        'port2': {'watts': 45, 'voltage': 15.0, 'current': 3},
    },
    4: {
        'label': '15W/15W',
        'port1': {'watts': 15, 'voltage': 5.0,  'current': 3},
        'port2': {'watts': 15, 'voltage': 5.0,  'current': 3},
    },
    5: {
        'label': '30W/30W',
        'port1': {'watts': 30, 'voltage': 15.0, 'current': 2},
        'port2': {'watts': 30, 'voltage': 15.0, 'current': 2},
    },
    6: {
        'label': '0W/60W',
        'port1': None,
        'port2': {'watts': 60, 'voltage': 20.0, 'current': 3},
    },
    7: {
        'label': '60W/0W',
        'port1': {'watts': 60, 'voltage': 20.0, 'current': 3},
        'port2': None,
    },
}

TOLERANCE = 0.25


def get_expected_voltage(mode, port):
    """
    Get the expected voltage for a given mode and port.

    Args:
        mode: Operating mode number (2-7)
        port: Port number (1 or 2)

    Returns:
        Expected voltage as float, or None if port is unused
    """
    mode_def = MODE_DEFINITIONS.get(mode)
    if mode_def is None:
        return None

    port_key = f'port{port}'
    port_def = mode_def.get(port_key)

    if port_def is None:
        return None

    return port_def['voltage']


def check_within_tolerance(mode, port, average):
    """
    Check if a voltage average is within tolerance for a given mode and port.

    Args:
        mode: Operating mode number
        port: Port number (1 or 2)
        average: Measured average voltage

    Returns:
        True if within tolerance, False if outside, None if port is unused
    """
    expected = get_expected_voltage(mode, port)

    if expected is None:
        return None

    lower = expected - TOLERANCE
    upper = expected + TOLERANCE

    return lower <= average <= upper


def check_reverse_polarity(results, device):
    """
    Check if all voltage averages for ONE device are within tolerance.

    Args:
        results: List of all result dictionaries
        device: Device number to check

    Returns:
        'Pass' if all within tolerance, 'Fail' if any outside
    """
    for r in results:
        if r['device'] != device:
            continue

        if not r.get('within_tolerance', True):
            return 'Fail'

    return 'Pass'


def process_files(filepaths, selected_modes=None):
    """
    Process a list of PassMark log files.
    Groups them by (device, mode, port), pairs angles, calculates averages.

    Args:
        filepaths: List of file paths
        selected_modes: List of mode numbers to process (e.g. [5, 6, 7])
                        If None, defaults to [5, 6, 7]

    Returns:
        Dictionary with results, warnings, errors, reverse_polarity
    """
    if selected_modes is None:
        selected_modes = [5, 6, 7]

    warnings = []
    errors = []

    # Step 1: Parse all filenames and get their averages
    file_data = []

    for fp in filepaths:
        file_info, file_warnings = parse_filename(fp)
        warnings.extend(file_warnings)

        if file_info is None:
            errors.append(f"Skipping file: {fp} (could not parse filename)")
            continue

        # Check if file's mode is in selected modes
        if file_info['mode'] not in selected_modes:
            warnings.append(
                f"File '{file_info['device']}-{file_info['mode']}-"
                f"{file_info['port']}-{file_info['angle']}.txt' is "
                f"Operating Mode {file_info['mode']} which is not selected "
                f"— file will be ignored"
            )
            continue

        avg = get_voltage_average(fp)

        if avg is None:
            errors.append(f"Warning: {fp} has no voltage data — recorded as 0")
            avg = 0.0
            
        file_data.append({
            'path': fp,
            'info': file_info,
            'average': avg
        })

    # Step 2: Group by (device, mode, port)
    groups = {}

    for fd in file_data:
        info = fd['info']
        key = (info['device'], info['mode'], info['port'])

        if key not in groups:
            groups[key] = {}

        angle = info['angle']

        if angle in groups[key]:
            warnings.append(
                f"Duplicate file for Device {info['device']}, "
                f"Operating Mode {info['mode']}, Port {info['port']}, "
                f"Angle {info['angle_degrees']}° — using the latest file"
            )

        groups[key][angle] = {
            'average': fd['average'],
            'angle_degrees': info['angle_degrees'],
            'path': fd['path']
        }

    # Step 3: Calculate results for each group
    results = []

    for key, angles in groups.items():
        device, mode, port = key

        if len(angles) == 2:
            avg1 = angles[1]['average']
            avg2 = angles[2]['average']
            combined = (avg1 + avg2) / 2
            combined_rounded = round(combined, 2)

            within = check_within_tolerance(mode, port, combined_rounded)

            results.append({
                'device': device,
                'mode': mode,
                'port': port,
                'paired': True,
                'average': combined_rounded,
                'angle': None,
                'within_tolerance': within if within is not None else True
            })

        elif len(angles) == 1:
            angle_key = list(angles.keys())[0]
            single = angles[angle_key]
            single_rounded = round(single['average'], 2)

            within = check_within_tolerance(mode, port, single_rounded)

            warnings.append(
                f"Device {device}, Operating Mode {mode}, Port {port}: "
                f"Only {single['angle_degrees']}° file provided"
            )

            results.append({
                'device': device,
                'mode': mode,
                'port': port,
                'paired': False,
                'average': single_rounded,
                'angle': single['angle_degrees'],
                'within_tolerance': within if within is not None else True
            })

    # Sort results by device, then mode, then port
    results.sort(key=lambda r: (r['device'], r['mode'], r['port']))

    # Step 4: Check reverse polarity per device
    devices = set(r['device'] for r in results)
    polarity = {}
    for device in devices:
        polarity[device] = check_reverse_polarity(results, device)

    return {
        'results': results,
        'warnings': warnings,
        'errors': errors,
        'reverse_polarity': polarity,
        'selected_modes': selected_modes,
        'mode_definitions': MODE_DEFINITIONS
    }