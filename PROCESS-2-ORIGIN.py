import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import originpro as op
import sys


# Code snippet 3: Create Origin graphs
def create_origin_graphs():
    def origin_shutdown_exception_hook(exctype, value, traceback):
        """Ensures Origin gets shut down if an uncaught exception"""
        op.exit()
        sys.__excepthook__(exctype, value, traceback)

    sys.excepthook = origin_shutdown_exception_hook

    # Only run if external Python
    if op.oext:
        # Create a new Origin project
        op.new()
        op.set_show(True)

    try:
        # Prompt user to select an Origin template file for graph 1
        template_file_path_1 = filedialog.askopenfilename(title="Select Origin Template File for Graph 1",
                                                          filetypes=[("Origin Template Files", "*.otpu")])

        # Prompt user to select a data file (CSV or XLSX) for the selected template
        data_file_types = [("CSV Files", "*.csv"), ("XLSX Files", "*.xlsx")]
        csv_file_path_1 = filedialog.askopenfilename(
            title=f"Select Data File for {os.path.basename(template_file_path_1)}",
            filetypes=data_file_types)

        # Load the CSV file into Origin worksheet
        wks = op.new_sheet()
        wks.from_file(csv_file_path_1, False)

        # Create a new graph using the selected template for graph 1
        gr1 = op.new_graph(template=template_file_path_1)
        gr1[0].add_plot(wks, 1, 0)
        gr1[0].rescale()

        # Add more graphs
        more_graphs = True
        graph_num = 2

        while more_graphs:
            # Prompt user to select an Origin template file for the graph
            template_file_path = filedialog.askopenfilename(title=f"Select Origin Template File for Graph {graph_num}",
                                                            filetypes=[("Origin Template Files", "*.otpu")])

            # Prompt user to select an XLSX file for the graph
            xlsx_file_path = filedialog.askopenfilename(
                title=f"Select XLSX File for {os.path.basename(template_file_path)}",
                filetypes=[("XLSX Files", "*.xlsx")])

            # Load the XLSX file into a new worksheet
            wks = op.new_sheet()
            wks.from_file(xlsx_file_path, False)

            # Create a graph page with the user-selected template for the graph
            gr = op.new_graph(template=template_file_path)
            gl = gr[0]

            # Prompt user for the range of columns to plot for the graph
            column_range = simpledialog.askstring("Range of Columns",
                                                  f"Enter the range of columns (e.g., 1-12) for Graph {graph_num}:")

            start_col, end_col = map(int, column_range.split("-"))
            y_columns = list(range(start_col, end_col + 1))

            for col in y_columns:
                plot = gl.add_plot(wks, col, 0)

            # Group and Rescale the graph
            gl.group()
            gl.rescale()

            # Prompt user if they want to add more graphs
            response = tk.messagebox.askyesno("Add More Graphs", "Do you want to add more graphs?")

            if response:
                graph_num += 1
            else:
                more_graphs = False

        # Tile all windows
        op.lt_exec('win-s T')

        # Prompt user for the output Origin file name
        output_file_path = filedialog.asksaveasfilename(title="Save Output Origin File",
                                                        filetypes=[("Origin Project Files", "*.opju")])

        # Creating a notes window
        nt = op.new_notes()

        # Appending input information to the notes
        nt.append("Input Information:")
        nt.append(f"Graph 1 Template: {os.path.basename(template_file_path_1)}")
        nt.append(f"Graph 1 CSV File: {os.path.basename(csv_file_path_1)}")

        for i in range(2, graph_num + 1):
            template_var = locals().get(f"template_file_path_{i}")
            csv_var = locals().get(f"csv_file_path_{i}")

            if template_var and csv_var:
                nt.append(f"\nGraph {i} Template: {os.path.basename(template_var)}")
                nt.append(f"Graph {i} CSV File: {os.path.basename(csv_var)}")

        # Appending folder path information to the notes
        nt.append("\nFolder Paths:")
        nt.append(f"Graph 1 Folder: {os.path.dirname(template_file_path_1)}")

        for i in range(2, graph_num + 1):
            template_var = locals().get(f"template_file_path_{i}")

            if template_var:
                nt.append(f"Graph {i} Folder: {os.path.dirname(template_var)}")

        nt.append(f"Output Folder: {os.path.dirname(output_file_path)}")

        # Appending output file information to the notes
        nt.append("\nOutput Information:")
        nt.append(f"Output File: {os.path.basename(output_file_path)}")

        # Displaying the note
        nt.view = 1

        # Tile all windows
        op.lt_exec('win-s T')

        # Save the project to the specified output file path
        if op.oext:
            output_file_path = os.path.abspath(output_file_path)
            op.save(output_file_path)

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        op.exit()


# End of Code snippet 3: Create Origin graphs


# Code snippet 4: Add graphs to existing Origin project
def add_graphs_to_project():
    def origin_shutdown_exception_hook(exctype, value, traceback):
        """Ensures Origin gets shut down if an uncaught exception"""
        op.exit()
        sys.__excepthook__(exctype, value, traceback)

    sys.excepthook = origin_shutdown_exception_hook

    # Only run if external Python
    if op.oext:
        op.set_show(True)

    def save_origin_project(output_path):
        try:
            op.save(output_path)
            return True  # Save operation successful
        except PermissionError:
            return False  # Save operation failed due to read-only
        except Exception as e:
            print(f"An error occurred while saving: {str(e)}")
            return False  # Save operation failed for other reasons

    def add_graphs_to_project_action():
        try:
            # Prompt user to select an existing Origin project file
            origin_file_path = filedialog.askopenfilename(title="Select Existing Origin Project File",
                                                          filetypes=[("Origin Project Files", "*.opju")])

            # Load the existing Origin project file
            op.open(origin_file_path)

            # Add more graphs
            more_graphs = True
            graph_num = 1

            while more_graphs:
                # Prompt user to select an Origin template file for the graph
                template_file_path = filedialog.askopenfilename(
                    title=f"Select Origin Template File for Graph {graph_num}",
                    filetypes=[("Origin Template Files", "*.otpu")])

                # Prompt user to select an XLSX file for the graph
                data_file_types = [("CSV Files", "*.csv"), ("XLSX Files", "*.xlsx")]
                xlsx_file_path = filedialog.askopenfilename(
                    title=f"Select XLSX File for {os.path.basename(template_file_path)}",
                    filetypes=data_file_types)

                # Load the XLSX file into a new worksheet
                wks = op.new_sheet()
                wks.from_file(xlsx_file_path, False)

                # Create a graph page with the user-selected template for the graph
                gr = op.new_graph(template=template_file_path)
                gl = gr[0]

                # Prompt user for the range of columns to plot for the graph
                column_range = simpledialog.askstring("Range of Columns",
                                                      f"Enter the range of columns (e.g., 1-12) for Graph {graph_num}:")

                start_col, end_col = map(int, column_range.split("-"))
                y_columns = list(range(start_col, end_col + 1))

                for col in y_columns:
                    gl.add_plot(wks, col, 0)

                # Group and Rescale the graph
                gl.group()
                gl.rescale()

                # Prompt user if they want to add more graphs
                response = tk.messagebox.askyesno("Add More Graphs", "Do you want to add more graphs?")

                if response:
                    graph_num += 1
                else:
                    more_graphs = False

            # Tile all windows
            op.lt_exec('win-s T')

            # Prompt user for the output Origin project file name and location
            output_file_path = filedialog.asksaveasfilename(title="Save Output Origin File As",
                                                            filetypes=[("Origin Project Files", "*.opju")])

            if not output_file_path:
                print("Save operation canceled by user.")
                op.exit()
                return  # Exit the function if the user cancels the save operation

            # Try to save the project file
            output_file_path = os.path.abspath(output_file_path)
            success = save_origin_project(output_file_path)

            # If the save operation failed due to read-only, prompt user for a new file path
            while not success:
                new_output_path = filedialog.asksaveasfilename(title="Save Output Origin File As",
                                                               filetypes=[("Origin Project Files", "*.opju")])
                if not new_output_path:
                    print("Save operation canceled by user.")
                    op.exit()
                    break  # Exit the loop if the user cancels the save operation

                new_output_path = os.path.abspath(new_output_path)
                success = save_origin_project(new_output_path)

                if success:
                    message = f"Processing completed! Output saved as:\n{new_output_path}"
                    messagebox.showinfo("Processing Complete", message)
                    op.exit()
                    break  # Exit the loop if the new save is successful
                else:
                    # Display an error message if the new save location is also read-only
                    messagebox.showerror("Error",
                                         "Selected save location is read-only. Please choose a different location.")
                    continue  # Continue the loop and prompt the user again

        except Exception as e:
            print(f"An error occurred: {str(e)}")
            op.exit()

    add_graphs_to_project_action()


root = tk.Tk()
root.withdraw()


# End of Code snippet 4: Add graphs to existing Origin project


# Code snippet 5: Function to exit the application
def exit_application():
    try:
        # Close the active Origin project
        op.exit()
        window.quit()  # Close the main GUI window, which ends the tkinter event loop
    except Exception as e:
        print(f"An error occurred while closing Origin: {str(e)}")


# End of Code snippet 5: Function to exit the application


# Create the main GUI window
window = tk.Tk()
window.title("FTIR Data Processing")
window.geometry("400x300")

# Create a frame for the header
header_frame = tk.Frame(window, padx=20, pady=20)
header_frame.pack()

# Create a label for the header
header_label = tk.Label(header_frame, text="FTIR Data Processing (part 2)", font=("Helvetica", 16, "bold"))
header_label.pack()

# Create a frame to contain the buttons and align it to the left
button_frame = tk.Frame(window, padx=20, pady=20)
button_frame.pack(anchor="w")  # Use "w" to anchor (justify) the frame to the left

# Create labels for step instructions
step3_label = tk.Label(button_frame, text="Step 3:")
step3_label.pack(anchor="w")  # Align label to the left

# Create a button for Step 3 functionality
create_origin_graphs_button = tk.Button(button_frame, text="Create Origin Project to add Graphs",
                                        command=create_origin_graphs, bg="light blue")
create_origin_graphs_button.pack(pady=5, anchor="w")

# Create a label for Step 4
step4_label = tk.Label(button_frame, text="Step 4:")
step4_label.pack(anchor="w")

# Create a button for Step 4 functionality
add_to_project_button = tk.Button(button_frame, text="Add Graphs to Existing Origin Project",
                                  command=add_graphs_to_project, bg="light blue")
add_to_project_button.pack(pady=5, anchor="w")

# Exit Section
exit_label = tk.Label(button_frame, text="To quit, click below")
exit_label.pack(anchor="w")

exit_button = tk.Button(button_frame, text="Exit Application", command=exit_application, bg="red")
exit_button.pack(pady=10, anchor="w")

# Start the tkinter event loop
window.mainloop()
