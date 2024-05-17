import tkinter as tk
from tkinter import filedialog
import fault_current as fc
import fault_voltage as fv
import fault_location as fl
import bus_finder as bf

# entry bus num
entry = 1
fault_type = ""

# Function to open file dialog and select Excel file
def select_excel_file():
    excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if excel_file:
        print("Selected Excel file:", excel_file)

def analyze_fault():
    substation_data = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    data = fl.extract_data_from_sheet(substation_data, 'Sheet1')
    fault_data = fl.filter_data_by_condition(data, -1, 'yes', 'fault_data')
    assigned_values = fl.assign_values(fault_data)
    fault_location = 0
    
    if fault_type == "3PF":
        fault_location = fl.calc_three_phase_fault_location(assigned_values)
    elif fault_type == "SLG":
        fault_location = fl.calc_phase_to_ground_fault_location(assigned_values)
    elif fault_type == "LLF":
        fault_location = fl.calc_phase_to_phase_fault_location(assigned_values)
    elif fault_type == "DLG":
        print("fault not available")

    fault_data = fl.extract_data_from_sheet(substation_data, 'Sheet1')
    time = fl.filter_data_by_condition(data, -1, 'yes', 'time')
    if isinstance(time, float):
        hours = int(time * 24)
        minutes = int(((time * 24) - hours) * 60)
        time = f"{hours}:{minutes:02d}"
    else:
        time = time
    
    new_window = tk.Tk()
    new_window.title("ALERT!")
    new_window.geometry("420x200")
    new_window.config(bg='#2b2b2b')
    label = tk.Label(new_window, text=
            f"ATTENTION: A FAULT has been DETECTED and LOCATED"
            f"\nwithin the system at {round(fault_location, 3)} miles" 
            f"\nat {time}. The fault type is {fault_type}."
            f"\nPlease take immediate ACTION to address the issue."
            "\nThank you."
            #f"\nThe system will just calculate the fault location at all possible senarios"
            f"\n\n\n\nA {fault_type} Fault is located at {round(fault_location, 4)} miles",
            bg='#2b2b2b', fg='white' ,font=("Default", 10))
    label.pack(pady=(40,0))
    
def calculate_fault_current():
      four_bus_system = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
      system_data = fc.extract_data_from_sheet(four_bus_system, 'Sheet1')
      I_3P = fc.calculate_3_phase_fault_current_(system_data=system_data, bus_num=entry)
      I_SLG = fc.calculate_SLG_fault_current_(system_data=system_data, bus_num=entry)
      I_LL = fc.calculate_LL_fault_current_(system_data=system_data, bus_num=entry)
      I_DLG = fc.calculate_DLG_fault_current_(system_data=system_data, bus_num=entry)
      new_window = tk.Tk()
      new_window.title("Fault Currents")
      new_window.geometry("500x330")
      new_window.config(bg='#2b2b2b')
      label = tk.Label(new_window, text= 
            "THE FAULT CURRENTS AT 3 PHASE FAULT ARE: "
            f"\nIa = {round(I_3P[0][0], 4)}cis({round(I_3P[0][1], 4)})deg" 
            f"\nIb = {round(I_3P[1][0], 4)}cis({round(I_3P[1][1], 4)})deg" 
            f"\nIc = {round(I_3P[2][0], 4)}cis({round(I_3P[2][1], 4)})deg"
            "\nTHE FAULT CURRENTS AT SLG FAULT ARE: "  
            f"\nIa = {round(I_SLG[0][0], 4)}cis({round(I_SLG[0][1], 4)})deg" 
            f"\nIb = {round(I_SLG[1][0], 4)}cis({round(I_SLG[1][1], 4)})deg" 
            f"\nIc = {round(I_SLG[2][0], 4)}cis({round(I_SLG[2][1], 4)})deg"
            "\nTHE FAULT CURRENTS AT LL FAULT ARE: " 
            f"\nIa = {round(I_LL[0][0], 4)}cis({round(I_LL[0][1], 4)})deg" 
            f"\nIb = {round(I_LL[1][0], 4)}cis({round(I_LL[1][1], 4)})deg" 
            f"\nIc = {round(I_LL[2][0], 4)}cis({round(I_LL[2][1], 4)})deg"
            "\nTHE FAULT CURRENTS AT DLG ARE: " 
            f"\nIa = {round(I_DLG[0][0], 4)}cis({round(I_DLG[0][1], 4)})deg" 
            f"\nIb = {round(I_DLG[1][0], 4)}cis({round(I_DLG[1][1], 4)})deg" 
            f"\nIc = {round(I_DLG[2][0], 4)}cis({round(I_DLG[2][1], 4)})deg"
            f"\n\nBUS {entry} FAULT", 
            bg='#2b2b2b', fg='white' ,font=("Default", 10))
      label.pack(pady=(20,0))

def calculate_fault_voltage():
      four_bus_system = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
      system_data = fv.extract_data_from_sheet(four_bus_system, 'Sheet1')
      V_3P = fv.calculate_3_phase_fault_voltage(system_data=system_data, bus_num=entry)
      V_SLG = fv.calculate_SLG_fault_voltage(system_data=system_data, bus_num=entry)
      V_LL = fv.calculate_LL_fault_voltage(system_data=system_data, bus_num=entry)
      V_DLG = fv.calculate_DLG_fault_voltage(system_data=system_data, bus_num=entry)
      new_window = tk.Tk()
      new_window.title("Fault Voltages")
      new_window.geometry("500x330")
      new_window.config(bg='#2b2b2b')
      label = tk.Label(new_window, text=
            "THE FAULT VOLTAGES AT 3 PHASE FAULT ARE: " 
            f"\nVa = {round(V_3P[0][0], 4)}cis({round(V_3P[0][1], 4)})deg" 
            f"\nVb = {round(V_3P[1][0], 4)}cis({round(V_3P[1][1], 4)})deg" 
            f"\nVc = {round(V_3P[2][0], 4)}cis({round(V_3P[2][1], 4)})deg"
            "\nTHE FAULT VOLTAGES AT SLG FAULT ARE: " 
            f"\nVa = {round(V_SLG[0][0], 4)}cis({round(V_SLG[0][1], 4)})deg" 
            f"\nVb = {round(V_SLG[1][0], 4)}cis({round(V_SLG[1][1], 4)})deg" 
            f"\nVc = {round(V_SLG[2][0], 4)}cis({round(V_SLG[2][1], 4)})deg"
            "\nTHE FAULT VOLTAGES AT LL FAULT ARE: " 
            f"\nVa = {round(V_LL[0][0], 4)}cis({round(V_LL[0][1], 4)})deg" 
            f"\nVb = {round(V_LL[1][0], 4)}cis({round(V_LL[1][1], 4)})deg" 
            f"\nVc = {round(V_LL[2][0], 4)}cis({round(V_LL[2][1], 4)})deg"
            "\nTHE FAULT VOLTAGES AT DLG FAULT ARE: " 
            f"\nVa = {round(V_DLG[0][0], 4)}cis({round(V_DLG[0][1], 4)})deg" 
            f"\nVb = {round(V_DLG[1][0], 4)}cis({round(V_DLG[1][1], 4)})deg" 
            f"\nVc = {round(V_DLG[2][0], 4)}cis({round(V_DLG[2][1], 4)})deg"
            f"\n\nBUS {entry} FAULT",
            bg='#2b2b2b', fg='white', font=("Default", 10))
      label.pack(pady=(20,0))
      
def calculate_bus_finder():
    row_num = entry
    file_data = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    fault_data = bf.assign_data(file_data, row_num, row_num)
    fault_current = [fault_data[i+1] for i in range(6)]
    fault_current = [[fault_current[i], fault_current[i+1]] for i in range(0, len(fault_current), 2)]
    fault_voltage = fault_data[7:13]
    fault_voltage = [[fault_voltage[i], fault_voltage[i+1]] for i in range(0, len(fault_voltage), 2)]
    file_path = "four-bus-power-system.xlsx"
    system_data = bf.extract_data_from_sheet(file_path, 'Sheet1', 2, 2)

    if fault_type == "3PF":
        BUS_FIND = bf.three_phase_fault(fault_current, fault_voltage, system_data)
    elif fault_type == "SLG":
        BUS_FIND = bf.SLG_fault(fault_current, fault_voltage, system_data)
    elif fault_type == "LLF":
        BUS_FIND = bf.LL_fault(fault_current, fault_voltage, system_data)
    elif fault_type == "DLG":
        BUS_FIND = bf.DLG_fault(fault_current, fault_voltage, system_data)

    print(BUS_FIND)
    print(fault_current)
    print(fault_voltage)


def main():
    # Create a Tkinter window
    root = tk.Tk()
    root.geometry("515x250")  # Width x Height
    root.title("FLS")
    root.config(bg='#2b2b2b') # Set the background color of the window
    # Add widgets
    label = tk.Label(root, text="Fault Location System", bg='#1a1a1a', fg='white')
    label.pack(pady=10)

    analyze_button = tk.Button(root, text="ANALYZE", bg='#2b2b2b', fg='white', command=analyze_fault, width=15, height=3)
    analyze_button.pack(side='left', padx=(40, 10), pady=(0, 150))

    info_label = tk.Label(root, text="Click on ANALYZE to analyze a system", bg="lightgrey")
    info_label.place(x=160, y=150)

    def show_selected():
        global fault_type
        fault_type = var.get()
        print("Selected option:", fault_type)
    var = tk.StringVar()
    radio_button1 = tk.Radiobutton(root, text="3PF  ", variable=var, value="3PF", command=show_selected)
    radio_button2 = tk.Radiobutton(root, text="SLG  ", variable=var, value="SLG", command=show_selected)
    radio_button3 = tk.Radiobutton(root, text="LLF  ", variable=var, value="LLF", command=show_selected)
    radio_button4 = tk.Radiobutton(root, text="DLG ", variable=var, value="DLG", command=show_selected)

    # Pack the radio buttons onto the window
    radio_button1.pack(side='left', padx=(40, 10), pady=(0, 150))
    radio_button2.pack(side='left', padx=(5, 10), pady=(0, 150))
    radio_button3.pack(side='left', padx=(5, 10), pady=(0, 150))
    radio_button4.pack(side='left', padx=(5, 10), pady=(0, 150))
    
    # Run the Tkinter event loop
    root.mainloop()

    
if __name__ == "__main__":
    main()
