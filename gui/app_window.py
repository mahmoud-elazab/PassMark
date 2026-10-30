"""
Main application window for PassMark Tool;
Uses customtkinter for a modern look.
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox

from core.calculator import process_files, MODE_DEFINITIONS
from core.filename_parser import parse_filename
from output.excel_writer import write_results_to_excel


class App(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        # Window Settings
        self.title("PassMark Tool")
        self.geometry("800x850")
        self.minsize(700, 750)

        # Store selected files and results
        self.selected_files = []
        self.processing_output = None
        self.reference_excel_path = None
        self.reference_currents = None


        self.stage_var = ctk.StringVar(value="Initial")  # default


        # Mode checkboxes variables
        # Default: 5, 6, 7 checked — 2, 3, 4 unchecked
        self.mode_vars = {
            2: ctk.BooleanVar(value=False),
            3: ctk.BooleanVar(value=False),
            4: ctk.BooleanVar(value=False),
            5: ctk.BooleanVar(value=True),
            6: ctk.BooleanVar(value=True),
            7: ctk.BooleanVar(value=True),
        }

        # Build the interface
        self.build_ui()

    def build_ui(self):
        """Create all the visual elements."""

        # Title Label
        self.title_label = ctk.CTkLabel(
            self,
            text="PassMark Tool",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.title_label.pack(pady=(15, 5))

        # Operating Mode(s) Selection Frame
        self.mode_frame = ctk.CTkFrame(self)
        self.mode_frame.pack(fill="x", padx=20, pady=(10, 5))

        # Label
        self.mode_label = ctk.CTkLabel(
            self.mode_frame,
            text="Operating Mode(s):",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.mode_label.grid(row=0, column=0, columnspan=3, padx=10, pady=(8, 4), sticky="w")

        # Row 1: Operating Mode 2, 3, 4
        modes_row1 = [2, 3, 4]
        for col, mode in enumerate(modes_row1):
            mode_def = MODE_DEFINITIONS[mode]
            cb = ctk.CTkCheckBox(
                self.mode_frame,
                text=f"Operating Mode {mode}  ({mode_def['label']})",
                variable=self.mode_vars[mode],
                font=ctk.CTkFont(size=12)
            )
            cb.grid(row=1, column=col, padx=15, pady=4, sticky="w")

        # Row 2: Operating Mode 5, 6, 7
        modes_row2 = [5, 6, 7]
        for col, mode in enumerate(modes_row2):
            mode_def = MODE_DEFINITIONS[mode]
            cb = ctk.CTkCheckBox(
                self.mode_frame,
                text=f"Operating Mode {mode}  ({mode_def['label']})",
                variable=self.mode_vars[mode],
                font=ctk.CTkFont(size=12)
            )
            cb.grid(row=2, column=col, padx=15, pady=4, sticky="w")

        # Make columns equal width so checkboxes align
        self.mode_frame.grid_columnconfigure(0, weight=1, uniform="mode_col")
        self.mode_frame.grid_columnconfigure(1, weight=1, uniform="mode_col")
        self.mode_frame.grid_columnconfigure(2, weight=1, uniform="mode_col")

        # Initial/Final section — separate frame below mode_frame
        self.stage_frame = ctk.CTkFrame(self)
        self.stage_frame.pack(fill="x", padx=20, pady=(5, 5))

        # Headline
        self.stage_label = ctk.CTkLabel(
            self.stage_frame,
            text="Initial/Final:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.stage_label.grid(
            row=0, column=0,
            padx=10, pady=(8, 4), sticky="w"
        )

        # Dropdown
        self.stage_optionmenu = ctk.CTkOptionMenu(
            self.stage_frame,
            values=["Initial", "Final"],
            variable=self.stage_var
        )
        self.stage_optionmenu.grid(
            row=1, column=0,
            padx=10, pady=(2, 10), sticky="w"
        )

        # Note about Operating Mode 1 stays in mode_frame
        self.mode_note = ctk.CTkLabel(
            self.mode_frame,
            text="ℹ: Operating Mode 1 is already included in the Excel sheet.",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.mode_note.grid(
            row=3, column=0, columnspan=3,
            padx=10, pady=(2, 8), sticky="w"
        )

        # Select Files Frame
        self.select_frame = ctk.CTkFrame(self)
        self.select_frame.pack(fill="x", padx=20, pady=(10, 5))

        self.select_button = ctk.CTkButton(
            self.select_frame,
            text="Select Files",
            command=self.on_select_files,
            width=150,
            height=35
        )
        self.select_button.pack(side="left", padx=10, pady=10)

        self.reference_button = ctk.CTkButton(
            self.select_frame,
            text="Load Reference Excel (optional)",
            command=self.on_select_reference_excel,
            width=220,
            height=35
        )
        self.reference_button.pack(side="left", padx=10, pady=10)

        self.reference_label = ctk.CTkLabel(
            self.select_frame,
            text="No reference Excel loaded",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.reference_label.pack(side="left", padx=10)

        self.file_count_label = ctk.CTkLabel(
            self.select_frame,
            text="No files selected",
            font=ctk.CTkFont(size=13)
        )
        self.file_count_label.pack(side="left", padx=10)

        # File List Label
        self.file_list_label = ctk.CTkLabel(
            self,
            text="Selected Files:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.file_list_label.pack(anchor="w", padx=20, pady=(10, 2))

        # File List Box
        self.file_listbox = ctk.CTkTextbox(self, height=100)
        self.file_listbox.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.file_listbox.configure(state="disabled")

        # Process Button
        self.process_button = ctk.CTkButton(
            self,
            text="▶ Process",
            command=self.on_process,
            width=150,
            height=35,
            state="disabled"
        )
        self.process_button.pack(pady=5)

        # Results Label
        self.results_label = ctk.CTkLabel(
            self,
            text="Results:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.results_label.pack(anchor="w", padx=20, pady=(10, 2))

        # Results Textbox (expandable)
        self.results_textbox = ctk.CTkTextbox(self)
        self.results_textbox.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.results_textbox.configure(state="disabled")

        # Text tags for coloring
        self.results_textbox._textbox.tag_config(
            "warning_header", foreground="red", font=("", 13, "bold")
        )
        self.results_textbox._textbox.tag_config(
            "warning_text", foreground="red", font=("", 12, "bold")
        )
        self.results_textbox._textbox.tag_config(
            "pass_normal", foreground="green"
        )

        self.results_textbox._textbox.tag_config(
            "fail_normal", foreground="red"
        )

        # Export Button
        self.export_button = ctk.CTkButton(
            self,
            text="Export to Excel",
            command=self.on_export,
            width=150,
            height=35,
            state="disabled"
        )
        self.export_button.pack(pady=(5, 15))

    def get_selected_modes(self):
        """Return list of selected mode numbers."""
        return [mode for mode, var in self.mode_vars.items() if var.get()]

    def validate_file_coverage(self):
        """
        Validate that selected modes have corresponding files, and that
        each mode has both ports and both angles per port when expected.

        Returns: list[str] warnings
        """
        warnings = []

        selected_modes = set(self.get_selected_modes())

        # Parse all files we have
        parsed = []
        for fp in self.selected_files:
            info, _ = parse_filename(fp)
            if info:
                parsed.append(info)

        # Group by mode
        files_by_mode = {}
        for info in parsed:
            m = info["mode"]
            files_by_mode.setdefault(m, []).append(info)

        # (1) Selected mode but no files
        for m in sorted(selected_modes):
            if m not in files_by_mode:
                warnings.append(
                    f"Operating Mode {m} is selected but no "
                    f"input files were provided for it."
                )

        # Only validate modes that have files AND are selected
        modes_to_check = sorted(set(files_by_mode.keys()) & selected_modes)

        expected_ports = {1, 2}
        expected_angles = {0, 180}

        for m in modes_to_check:
            infos = files_by_mode[m]

            # (2) Port completeness
            ports_present = {i["port"] for i in infos}
            missing_ports = expected_ports - ports_present
            if missing_ports:
                warnings.append(
                    f"Operating Mode {m}: missing files for "
                    f"Port(s) {sorted(missing_ports)}."
                )

            # (3) Angle completeness per port
            infos_by_port = {}
            for i in infos:
                infos_by_port.setdefault(i["port"], []).append(i)

            for port, port_infos in infos_by_port.items():
                angles_present = {
                    pi.get("angle_degrees")
                    for pi in port_infos
                }
                angles_present.discard(None)

                missing_angles = expected_angles - angles_present
                if missing_angles and angles_present:
                    warnings.append(
                        f"Operating Mode {m}, Port {port}: "
                        f"missing angle(s) {sorted(missing_angles)}° "
                        f"(only {sorted(angles_present)}° found)."
                    )

        return warnings

    def on_select_reference_excel(self):
        """Load optional reference Excel and parse unloaded/loaded currents."""
        fp = filedialog.askopenfilename(
            title="Select Reference Excel (optional)",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not fp:
            return

        try:
            from core.reference_excel_reader import read_reference_currents
            import os

            self.reference_excel_path = fp
            self.reference_currents = read_reference_currents(fp)

            self.reference_label.configure(
                text=f"Reference loaded: {os.path.basename(fp)}",
                text_color="green"
            )

        except Exception as e:
            self.reference_excel_path = None
            self.reference_currents = None
            self.reference_label.configure(
                text="Failed to load reference Excel",
                text_color="red"
            )
            messagebox.showerror("Reference Excel Error", str(e))

    


    def on_select_files(self):
        """Handle the Select Files button click."""

        files = filedialog.askopenfilenames(
            title="Select PassMark Log Files",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not files:
            return

        self.selected_files = list(files)

        # Update file count label
        count = len(self.selected_files)
        self.file_count_label.configure(text=f"{count} file(s) selected ✅")

        # Update file list display
        self.file_listbox.configure(state="normal")
        self.file_listbox.delete("1.0", "end")

        selected_modes = self.get_selected_modes()

        for fp in self.selected_files:
            file_info, warnings = parse_filename(fp)
            if file_info:
                # Warn if file mode not in selected modes
                mode_note = ""
                if file_info['mode'] not in selected_modes:
                    mode_note = "  ⚠ Mode not selected"

                line = (
                    f"{file_info['device']}-{file_info['mode']}-"
                    f"{file_info['port']}-{file_info['angle']}  "
                    f"(Operating Mode {file_info['mode']}, "
                    f"Port {file_info['port']}, "
                    f"{file_info['angle_degrees']}°)"
                    f"{mode_note}\n"
                )
            else:
                line = f"⚠️ Unrecognized filename format\n"

            self.file_listbox.insert("end", line)

        self.file_listbox.configure(state="disabled")

        # Enable Process button
        self.process_button.configure(state="normal")

        # Reset results
        self.results_textbox.configure(state="normal")
        self.results_textbox.delete("1.0", "end")
        self.results_textbox.configure(state="disabled")
        self.export_button.configure(state="disabled")

    def on_process(self):
        """Handle the Process button click."""

        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select files first.")
            return

        # Get selected modes
        selected_modes = self.get_selected_modes()

        if not selected_modes:
            messagebox.showwarning(
                "No Modes Selected",
                "Please select at least one Operating Mode."
            )
            return

        # Pre-check: validate file coverage
        pre_warnings = self.validate_file_coverage()

        # Process the files
        self.processing_output = process_files(
            self.selected_files,
            selected_modes=selected_modes
        )

        results = self.processing_output['results']
        warnings = self.processing_output['warnings']
        errors = self.processing_output['errors']

        # Display everything in one block
        self.results_textbox.configure(state="normal")
        self.results_textbox.delete("1.0", "end")

        # Coverage warnings (mode/port/angle)
        if pre_warnings:
            self.results_textbox.insert(
                "end", "Coverage Warnings:\n", "warning_header"
            )
            for w in pre_warnings:
                self.results_textbox.insert(
                    "end", f"  {w}\n", "warning_text"
                )
            self.results_textbox.insert("end", "\n")

        # Errors
        if errors:
            self.results_textbox.insert(
                "end", "Errors:\n", "warning_header"
            )
            for e in errors:
                self.results_textbox.insert(
                    "end", f"  {e}\n", "warning_text"
                )
            self.results_textbox.insert("end", "\n")

        # Warnings from calculator
        if warnings:
            self.results_textbox.insert(
                "end", "Warnings:\n", "warning_header"
            )
            for w in warnings:
                self.results_textbox.insert(
                    "end", f"  {w}\n", "warning_text"
                )
            self.results_textbox.insert("end", "\n")

        # Results
        if results:
            current_device = None
            polarity = self.processing_output.get('reverse_polarity', {})

            for r in results:
                if r['device'] != current_device:
                    current_device = r['device']
                    device_polarity = polarity.get(current_device, '')
                    self.results_textbox.insert(
                        "end",
                        f"\n  Device {current_device}  —  Reverse Polarity: "
                    )
                    if device_polarity == 'Pass':
                        self.results_textbox.insert(
                            "end", "Pass ✅\n", "pass_normal"
                        )
                    else:
                        self.results_textbox.insert(
                            "end", "Fail ❌\n", "fail_normal"
                        )

                mode_def = MODE_DEFINITIONS.get(r['mode'], {})
                mode_label = mode_def.get('label', '')

                if r['paired']:
                    line = (
                        f"    Operating Mode {r['mode']} ({mode_label}), "
                        f"Port {r['port']}  →  {r['average']} V  "
                    )
                else:
                    line = (
                        f"    Operating Mode {r['mode']} ({mode_label}), "
                        f"Port {r['port']} ({r['angle']}° only)  →  "
                        f"{r['average']} V  "
                    )

                self.results_textbox.insert("end", line)
                if r.get('within_tolerance', True):
                    self.results_textbox.insert("end", "✅\n", "pass_normal")
                else:
                    self.results_textbox.insert("end", "❌\n", "fail_normal")

            self.export_button.configure(state="normal")

        else:
            self.results_textbox.insert(
                "end", "No results to display.\n"
            )

        self.results_textbox.configure(state="disabled")


    def on_export(self):
        """Handle the Export to Excel button click."""

        if self.processing_output is None:
            messagebox.showwarning("No Results", "Please process files first.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="voltage_results.xlsx"
        )

        if not filepath:
            return

        try:
            # Dropdown selection: "Initial" or "Final"
            stage_label = self.stage_var.get()

            write_results_to_excel(
                self.processing_output,
                filepath,
                stage_label=stage_label,
                reference_currents=self.reference_currents
            )
            messagebox.showinfo("Success", f"Excel file saved!\n{filepath}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file:\n{e}")