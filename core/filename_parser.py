"""
Parses filenames to extract test metadata.

Filename format: {device}-{mode}-{port}-{angle}.txt
Example: 379353-5-1-1.txt
"""
import os
import re

# Known valid values (used for warnings, NOT for blocking)
KNOWN_MODES = {0, 1, 2, 3, 4, 5, 6, 7}
KNOWN_PORTS = {1, 2}
ANGLE_MAP = {
    1: 0,
    2: 180
}


def parse_filename(filepath):
    """
    Extract device number, mode, port, and angle from a filename.

    Args:
        filepath: Full path or just filename like "379353-5-1-1.txt"

    Returns:
        Tuple of (file_info dict, warnings list)
        file_info is None if filename doesn't match the general pattern.
        warnings is a list of strings (empty if everything looks normal).
    """
    filename = os.path.basename(filepath)
    name_without_ext = os.path.splitext(filename)[0]

    pattern = r'^(\d+)-(\d+)-(\d+)-(\d+)$'
    match = re.match(pattern, name_without_ext)

    if not match:
        return None, [f"Filename '{filename}' does not match expected pattern."]

    device = match.group(1)
    mode = int(match.group(2))
    port = int(match.group(3))
    angle = int(match.group(4))

    # Collect warnings for unusual values
    warnings = []

    if mode not in KNOWN_MODES:
        warnings.append(f"Mode {mode} is unusual (expected 0-7)")

    if port not in KNOWN_PORTS:
        warnings.append(f"Port {port} is unusual (expected 1 or 2)")

    if angle not in ANGLE_MAP:
        warnings.append(f"Angle {angle} is unknown (expected 1 or 2)")

    angle_degrees = ANGLE_MAP.get(angle, f"unknown({angle})")

    file_info = {
        'device': device,
        'mode': mode,
        'port': port,
        'angle': angle,
        'angle_degrees': angle_degrees
    }

    return file_info, warnings