import pandas as pd
import os
import sys

def read_capacities(file_name):
    with open(file_name, 'r') as file:
        capacities = [float(line.strip()) for line in file]
    return capacities

def write_capacities(file_name, capacities):
    with open(file_name, 'w') as file:
        for capacity in capacities:
            file.write(str(capacity) + '\n')

def assemble_battery_pack(capacities, num_series, num_parallel):
    if num_series * num_parallel > len(capacities):
        return None

    capacities.sort(reverse=True)
    battery_pack = [[] for _ in range(num_series)]
    
    for _ in range(num_parallel):
        for s in battery_pack:
            if capacities:
                s.append(capacities.pop(0))
            else:
                break
    
    return battery_pack

def calculate_voltages(num_series, chemistry):
    if chemistry.lower() == 'lion':
        cutoff_voltage_base = 2.8
        nominal_voltage_base = 3.7
        fully_charged_voltage_base = 4.2
    elif chemistry.lower() == 'lifepo4':
        cutoff_voltage_base = 2
        nominal_voltage_base = 3.2
        fully_charged_voltage_base = 3.6
    else:
        print("Invalid battery chemistry.")
        return None, None, None

    cutoff_voltage = cutoff_voltage_base * num_series
    nominal_voltage = nominal_voltage_base * num_series
    fully_charged_voltage = fully_charged_voltage_base * num_series

    return cutoff_voltage_base, cutoff_voltage, nominal_voltage_base, nominal_voltage, fully_charged_voltage_base, fully_charged_voltage

def get_integer_input(prompt):
    while True:
        try:
            value = int(input(prompt))
            return value
        except ValueError:
            print("Please enter a valid number.")

def main():
    file_name = "capacities.txt"
    if not os.path.exists(file_name):
        print("Capacities.txt file is missing. Please press enter to quit")
        input()
        sys.exit()

    capacities = read_capacities(file_name)
    
    while True:
        num_series = get_integer_input("Enter the number of series: ")
        num_parallel = get_integer_input("Enter the number of parallel cells: ")
        
        battery_chemistry = ""
        while battery_chemistry.lower() not in ['lion', 'lifepo4']:
            battery_chemistry = input("Enter the battery chemistry (Lion/LiFePo4): ")
        
        num_packs = get_integer_input("Enter the number of packs needed: ")
        
        output_option = ""
        while output_option.lower() not in ['terminal', 'excel']:
            output_option = input("Do you want to output to terminal or save to an Excel file? (terminal/excel): ")
        
        if output_option.lower() == 'excel':
            excel_file_name = input("Enter the Excel file name: ")
            if not excel_file_name.endswith('.xlsx'):
                excel_file_name += '.xlsx'
            while os.path.exists(excel_file_name):
                overwrite = input(f"The file {excel_file_name} already exists. Overwrite existing file? (yes/no): ")
                if overwrite.lower() == 'yes':
                    break
                else:
                    excel_file_name = input("Please choose a different file name: ")
                    if not excel_file_name.endswith('.xlsx'):
                        excel_file_name += '.xlsx'
            try:
                writer = pd.ExcelWriter(excel_file_name, engine='openpyxl')
            except PermissionError:
                print("Error saving file. Please check it is not opened by another program or you have permission to save in the location.")
                return
        
        for pack_num in range(1, num_packs + 1):
            print(f"\nPack {pack_num}:")
            battery_pack = assemble_battery_pack(capacities, num_series, num_parallel)
            if battery_pack is not None:
                total_pack_capacity = 0
                series_capacities = []
                for i, series in enumerate(battery_pack, start=1):
                    total_series_capacity = sum(series)
                    total_pack_capacity += total_series_capacity
                    series_capacities.append(total_series_capacity)
                    if output_option.lower() == 'terminal':
                        print(f"Series {i}: {series}, Total Series Capacity: {total_series_capacity}mAh")
                
                total_pack_capacity = round(total_pack_capacity / num_series, 2)
                max_series_cell_diff = round((max(series_capacities) - min(series_capacities)) / max(series_capacities) * 100, 2)
                if max_series_cell_diff > 5:
                    print("Warning: The percent difference between the highest and lowest capacity series cell is over 5%.")
                cutoff_voltage_base, cutoff_voltage, nominal_voltage_base, nominal_voltage, fully_charged_voltage_base, fully_charged_voltage = calculate_voltages(num_series, battery_chemistry)
                if output_option.lower() == 'terminal':
                    print(f"Total Pack Capacity: {total_pack_capacity}mAh, Max series cell difference: {max_series_cell_diff}%")
                    print(f"Cut Off Voltage ({cutoff_voltage_base}): {cutoff_voltage}V")
                    print(f"Nominal Voltage ({nominal_voltage_base}): {nominal_voltage}V")
                    print(f"Fully Charged Voltage ({fully_charged_voltage_base}): {fully_charged_voltage}V")
                elif output_option.lower() == 'excel':
                    df = pd.DataFrame(battery_pack, columns=[f'Parallel Cell {i+1}' for i in range(num_parallel)], index=[f'Series {i+1}' for i in range(num_series)])
                    df['Total Series Capacity'] = [f'{capacity}mAh' for capacity in series_capacities]
                    df.loc['Total Pack Capacity'] = ['' for _ in range(num_parallel)] + [f"{total_pack_capacity}mAh, Max series cell difference: {max_series_cell_diff}%"]
                    df.loc['Cut Off Voltage'] = ['' for _ in range(num_parallel)] + [f"({cutoff_voltage_base}) {cutoff_voltage}V"]
                    df.loc['Nominal Voltage'] = ['' for _ in range(num_parallel)] + [f"({nominal_voltage_base}) {nominal_voltage}V"]
                    df.loc['Fully Charged Voltage'] = ['' for _ in range(num_parallel)] + [f"({fully_charged_voltage_base}) {fully_charged_voltage}V"]
                    df.to_excel(writer, sheet_name=f'Pack {pack_num}')
                print("Complete")
            else:
                print(f"Not enough cells left to create another {num_series}s{num_parallel}p pack!")
                break

        if output_option.lower() == 'excel':
            writer._save()

        remove_used = ""
        while remove_used.lower() not in ['yes', 'no']:
            remove_used = input("\nDo you want to remove the used cells from the file? (yes/no): ")
        if remove_used.lower() == 'yes':
            write_capacities(file_name, capacities)

        make_another_pack = ""
        while make_another_pack.lower() not in ['yes', 'no']:
            make_another_pack = input("\nDo you want to make another pack? (yes/no): ")
        if make_another_pack.lower() == 'no':
            break

if __name__ == "__main__":
    main()
