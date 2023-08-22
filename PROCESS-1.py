import os
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox
import pandas as pd


# Code snippet 1: Rename columns
def rename_columns():
    def rename_columns_action():
        # Select the CSV file
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])

        if file_path:
            try:
                # Prompt for start and end voltage values
                start_voltage = float(simpledialog.askstring("Start Voltage", "Enter the starting voltage:"))
                mid_voltage = float(simpledialog.askstring("Mid Voltage", "Enter the ending voltage:"))

                # Read the CSV file
                df = pd.read_csv(file_path)

                # Get the headers
                headers = df.columns.tolist()

                # Calculate the voltage step
                step_up = (mid_voltage - start_voltage) / (len(headers) // 2 - 1)
                step_down = (start_voltage - mid_voltage) / (len(headers) // 2 - 1)

                # Rename the columns
                for i in range(1, len(headers)):
                    if i <= len(headers) // 2:
                        voltage = start_voltage + (i - 1) * step_up
                    else:
                        voltage = mid_voltage + (i - len(headers) // 2 - 1) * step_down
                    new_header = "{:.2f}".format(voltage)  # Format the voltage with 2 decimal places
                    df.rename(columns={headers[i]: new_header}, inplace=True)

                # Save the renamed data to XLSX file
                save_path = os.path.splitext(file_path)[0] + "_renamed.xlsx"
                df.to_excel(save_path, index=False)

                # Open the folder where the output files are saved
                os.startfile(save_path)

                messagebox.showinfo("Success", "Columns renamed successfully. Saved as " + save_path)
            except Exception as e:
                messagebox.showerror("Error", "An error occurred: " + str(e))

    rename_columns_action()


root = tk.Tk()
root.withdraw()


# End of Code snippet 1: Rename columns

# code 2 start
def bg_processing():
    # Select the CSV file
    xlsx_file_path = filedialog.askopenfilename(title="Select Input XLSX File", filetypes=[("XLSX Files", "*.xlsx")])

    if xlsx_file_path:
        try:
            # Read the input XLSX file and rename the sheets
            xlsx = pd.ExcelFile(xlsx_file_path)
            sheets = xlsx.sheet_names
            df = None  # Initialize df variable
            wavenumber_column = None  # Initialize wavenumber_column variable

            for i, sheet in enumerate(sheets):
                df_sheet = pd.read_excel(xlsx_file_path, sheet_name=sheet)
                df_sheet.to_excel(xlsx_file_path, sheet_name=f"Sheet{i + 1}", index=False)

                if i == 0:
                    df = df_sheet
                    wavenumber_column = df_sheet["Wavenumber"]

            # Create a Tkinter window for column selection
            column_window = tk.Tk()
            column_window.title("Select Column")
            column_window.geometry("300x100")

            # Create a label for column selection
            column_label = ttk.Label(column_window, text="Choose a column:")
            column_label.pack()

            # Create a combobox for column selection
            column_combobox = ttk.Combobox(column_window, values=df.columns.tolist())
            column_combobox.pack()

            # Create a button to confirm column selection
            confirm_button = ttk.Button(column_window, text="Confirm", command=column_window.quit)
            confirm_button.pack()

            # Run the column selection window
            column_window.mainloop()
            # Get the chosen column
            chosen_column = column_combobox.get()

            # Create a new sheet for processing
            processed_sheet = pd.DataFrame()
            processed_sheet["Wavenumber"] = wavenumber_column

            for column in df.columns[1:]:
                if column == chosen_column:
                    processed_sheet[column] = 0
                else:
                    processed_sheet[column] = df[column] - df[chosen_column]

            # Prompt for the output directory
            xlsx_output_dir = filedialog.askdirectory(title="Select Output Directory")

            # Get the input filename without extension
            input_xlsx_file_name = os.path.splitext(os.path.basename(xlsx_file_path))[0]

            # Construct the output file path
            output_xlsx_file_name = f"{input_xlsx_file_name}_{chosen_column}.xlsx"
            output_xlsx_file_path = os.path.join(xlsx_output_dir, output_xlsx_file_name)

            # Save the processed sheet to a new workbook
            with pd.ExcelWriter(output_xlsx_file_path, engine="openpyxl") as writer:
                processed_sheet.to_excel(writer, sheet_name="Sheet1", index=False)

            # Open the folder where the output files are saved
            os.startfile(xlsx_output_dir)

            # Show completion message
            message = f"Processing completed! Output saved as:\n{output_xlsx_file_path}"
            messagebox.showinfo("Processing Complete", message)
            column_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", "An error occurred: " + str(e))


root = tk.Tk()
root.withdraw()


# Code snippet 2 end

# Code snippet 5: Function to exit the application
def exit_application():
    try:
        window.quit()  # Close the main GUI window, which ends the tkinter event loop
    except Exception as e:
        print(f"An error occurred while closing Origin: {str(e)}")


# End of Code snippet 5: Function to exit the application


# Create the main GUI window
window = tk.Tk()
window.title("FTIR Data Processing")
window.geometry("300x230")

# Create a frame to contain the buttons and align it to the left
button_frame = tk.Frame(window, padx=20, pady=20)
button_frame.pack(anchor="w")  # Use "w" to anchor (justify) the frame to the left

# Create buttons with some styling
button_style = {
    "font": ("Helvetica", 14),
    "width": 30,
    "height": 2,
    "bd": 2,  # Border thickness
    "relief": "solid",  # Border style
}

# Create buttons for each functionality
rename_columns_button = tk.Button(window, text="Step 1: Rename Columns", command=rename_columns, bg="light blue")
process_background_data_button = tk.Button(window, text="Step 2: Reprocess Background", command=bg_processing,
                                           bg="light blue")
exit_button = tk.Button(window, text="Exit Application", command=exit_application, bg="red")

# Pack the buttons with some spacing and justify to the left
rename_columns_button.pack(pady=10, anchor="w")
process_background_data_button.pack(pady=10, anchor="w")
exit_button.pack(pady=10, anchor="w", )

# Start the tkinter event loop
window.mainloop()
