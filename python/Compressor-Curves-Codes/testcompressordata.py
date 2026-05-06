# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import win32com.client

class testcompressordata:
    def __init__(self, unisim_path):
        """Initialize UniSim connection but don't execute anything yet."""
        self.unisim_path = unisim_path
        self.unisim_app = None
        self.case = None

    def connect(self):
        """Establish UniSim connection (only when called)."""
        if not self.unisim_app:
            self.unisim_app = win32com.client.Dispatch("UniSimDesign.Application")
            self.unisim_app.Visible = True
            self.case = self.unisim_app.SimulationCases.Open(self.unisim_path)

            if not self.case:
                raise RuntimeError("Error: Failed to load UniSim simulation case!")

    def __getitem__(self, keys):
        """Allows dictionary-style access: extract_data[subflowsheet] or extract_data[subflowsheet, compressor]."""
        self.connect()  # Ensure connection before accessing data

        if isinstance(keys, tuple) and len(keys) == 2:
            subflowsheet_tag, compressor_name = keys
            return self.get_compressor_head_flow(subflowsheet_tag, compressor_name)
        elif isinstance(keys, str):
            return self.get_operations(keys)
        else:
            raise ValueError("Invalid input format. Use extract_data['Subflowsheet'] or extract_data['Subflowsheet', 'Compressor'].")

    def get_operations(self, subflowsheet_tag):
        """Retrieve all operations within the subflowsheet."""
        self.connect()
        flowsheet = self.case.Flowsheet.Flowsheets.Item(subflowsheet_tag)
        return [op.Name for op in flowsheet.Operations]

    def get_compressor_head_flow(self, subflowsheet_tag, compressor_name):
        """Extract head (m), flow (ACT m³/h), and indexed curve names."""
        self.connect()
        flowsheet = self.case.Flowsheet.Flowsheets.Item(subflowsheet_tag)
        compressor = flowsheet.Operations.Item(compressor_name)

        head_values = []
        flow_values_sec = []
        curve_names = []

        for curve_index in range(compressor.Curves.Count):
            curve = compressor.Curves.Item(curve_index)
            curve_names.append(f"{curve_index + 1}: {curve.Name}")  # Indexed curve name
            head_values.append([value * 9.81 / 1000 for value in curve.HeadValue])  ## ERIN
            flow_values_sec.append(curve.GasFlowRateValue)

        # Convert flow from ACT m³/s to ACT m³/h
        flow_values_hr = [[flow * 3600 for flow in curve] for curve in flow_values_sec]

        return curve_names, head_values, flow_values_hr

    def get_compressor_head_flow_main(self, compressor_name):
        """Extract head (m), flow (ACT m³/h), and indexed curve names from the main flowsheet."""
        self.connect()
        flowsheet = self.case.Flowsheet  # Main flowsheet
        compressor = flowsheet.Operations.Item(compressor_name)

        head_values = []
        flow_values_sec = []
        curve_names = []

        for curve_index in range(compressor.Curves.Count):
            curve = compressor.Curves.Item(curve_index)
            curve_names.append(f"{curve_index + 1}: {curve.Name}")  # Indexed curve name
            head_values.append([value * 9.81 / 1000 for value in curve.HeadValue]) ## ERIN
            flow_values_sec.append(curve.GasFlowRateValue)

        # Convert flow from ACT m³/s to ACT m³/h
        flow_values_hr = [[flow * 3600 for flow in curve] for curve in flow_values_sec]

        return curve_names, head_values, flow_values_hr
    


    def display_results(self, *args):
        """Print head & flow values for compressors in both main flowsheet and subflowsheets."""
        self.connect()

        if len(args) == 1:
            compressor_name = args[0]
            curve_names, head_values, flow_values_hr = self.get_compressor_head_flow_main(compressor_name)
        elif len(args) == 2:
            subflowsheet_tag, compressor_name = args
            curve_names, head_values, flow_values_hr = self.get_compressor_head_flow(subflowsheet_tag, compressor_name)
        else:
            raise ValueError("Invalid input format. Use extract_data.display_results('Compressor') or extract_data.display_results('Subflowsheet', 'Compressor').")

        if not curve_names:
            print(f"Error: No curve data found for {compressor_name}")
            return
        

        print("\n--- Compressor Head & Flow Data ---")
        for index, (name, head, flow) in enumerate(zip(curve_names, head_values, flow_values_hr), start=1):
            head_str = ', '.join(map(str, head))  # Convert head list to a single string
            flow_str = ', '.join(map(str, flow))  # Convert flow list to a single string
            print(f"{name}: Head (m) = [{head_str}], Flow (ACT m³/h) = [{flow_str}]\n\n")
        
        return head_values, flow_values_hr ## ERIN
