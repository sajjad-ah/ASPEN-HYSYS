# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import win32com.client

class ResultPressureReduction:
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

    def find_stream_in_subflowsheet(self, subflowsheet_name, stream_name):
        """Search for a stream inside a subflowsheet."""
        self.connect()
        flowsheet = self.case.Flowsheet.Flowsheets.Item(subflowsheet_name)

        print(f"\nSearching for stream '{stream_name}' inside subflowsheet '{subflowsheet_name}'...")
        for stream in flowsheet.Streams:
            print("Checking stream:", stream.Name)  # Debugging output
            if stream.Name == stream_name:
                print(f"Stream '{stream_name}' found!")
                return stream  # Return stream object if found

        print(f"Error: Stream '{stream_name}' not found in subflowsheet '{subflowsheet_name}'.")
        return None  # Return None if not found

    def list_stream_parameters(self, subflowsheet_name, stream_name):
        """List available parameters for a given stream inside a subflowsheet."""
        self.connect()
        stream = self.find_stream_in_subflowsheet(subflowsheet_name, stream_name)

        if not stream:
            print(f"Error: Stream '{stream_name}' not found.")
            return

        print(f"\nAvailable parameters for stream '{stream_name}':")
        for param in stream.Parameters:
            print(param.Name)  # Show all possible properties

    def get_stream_properties(self, subflowsheet_name, stream_name, properties):
        """Retrieve specified properties (e.g., flowrate, density, pressure) for a given stream inside a subflowsheet."""
        self.connect()
        stream = self.find_stream_in_subflowsheet(subflowsheet_name, stream_name)

        if not stream:
            print(f"Error: Stream '{stream_name}' not found.")
            return None

        results = {}
        for prop in properties:
            try:
                results[prop] = stream.Parameters.Item(prop).Value  # Correct property retrieval
            except AttributeError:
                print(f"Warning: Property '{prop}' not found in stream '{stream_name}'")
                results[prop] = None  # Return None if property doesn't exist

        return results

    def display_stream_data(self):
        """Retrieve and display flowrate, density, and pressure for the specified streams inside subflowsheet."""
        self.connect()

        # Stream "20L8025" (Flowrate & Density)
        stream_20L8025_data = self.get_stream_properties("TPL16", "20L8025", ["MassFlow", "MassDensity"])
        
        if stream_20L8025_data:
            flowrate_kg_hr = stream_20L8025_data.get("MassFlow", "N/A")  
            density_kg_m3 = stream_20L8025_data.get("MassDensity", "N/A")
            print(f"Stream '20L8025' - Flowrate: {flowrate_kg_hr} kg/hr, Density: {density_kg_m3} kg/m³")
        else:
            print("Error: Failed to retrieve '20L8025' stream properties.")

        # Stream "Manifold B_1" (Pressure)
        stream_manifold_data = self.get_stream_properties("TPL16", "Manifold B_1", ["Pressure"])
        
        if stream_manifold_data:
            pressure_barg = stream_manifold_data.get("Pressure", "N/A")
            print(f"Stream 'Manifold B_1' - Pressure: {pressure_barg} barg")
        else:
            print("Error: Failed to retrieve 'Manifold B_1' stream properties.")
            
def list_stream_attributes(self, subflowsheet_name, stream_name):
    """List all attributes available for a given stream."""
    self.connect()
    stream = self.find_stream_in_subflowsheet(subflowsheet_name, stream_name)

    if not stream:
        print(f"Error: Stream '{stream_name}' not found.")
        return

    print(f"\nAvailable attributes for stream '{stream_name}':")
    for attr in dir(stream):  # Lists all attributes
        print(attr)

# Example Usage
# unisim_path = r"C:\Users\sahm\Downloads\1_Sugg\PRconsequences_MASTER high export.usc"
# extract_data = ResultPressureReduction(unisim_path)
# extract_data.list_stream_parameters("TPL16", "20L8025")  # Lists available properties
# extract_data.display_stream_data()  # Retrieves and displays properties
