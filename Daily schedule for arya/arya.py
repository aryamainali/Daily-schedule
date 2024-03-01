import tkinter as tk
from tkinter import messagebox, ttk, filedialog 
import pandas as pd
import os

class ScheduleApp:
    def __init__(self, master):
        # Initialize the application window
        self.master = master
        self.master.title("Daily Schedule App")
        
        # Initialize an empty dictionary to store schedule data
        self.schedule = {}

        # Create the GUI widgets
        self.create_widgets()

    def create_widgets(self):
        # Styling for GUI widgets
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")  # Background color for frame
        self.style.configure("TButton", background="#007bff", foreground="black", font=("Arial", 12), padding=(10, 5))  # Styling for buttons
        self.style.map("TButton", background=[("active", "pink")])  # Styling for button when active
        self.style.configure("TLabel", background="lightblue", font=("Arial", 12))  # Styling for labels
        self.style.configure("TEntry", background="lightblue", font=("Arial", 12))  # Styling for entry fields

        # Create frame
        self.frame = ttk.Frame(self.master)
        self.frame.grid(row=0, column=0)

        # Labels and entry fields for task and time
        self.task_label = ttk.Label(self.frame, text="Task:")
        self.task_label.grid(row=0, column=0, padx=10, pady=10)
        self.task_entry = ttk.Entry(self.frame, width=40)
        self.task_entry.grid(row=0, column=1, padx=10, pady=10)

        self.time_label = ttk.Label(self.frame, text="Time:")
        self.time_label.grid(row=1, column=0, padx=10, pady=10)
        self.time_entry = ttk.Entry(self.frame, width=40)
        self.time_entry.grid(row=1, column=1, padx=10, pady=10)

        # Buttons for adding, displaying, removing, updating, and exporting schedule
        self.add_button = ttk.Button(self.frame, text="Add Task", command=self.add_task)
        self.add_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.display_button = ttk.Button(self.frame, text="Display Schedule", command=self.display_schedule)
        self.display_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.remove_button = ttk.Button(self.frame, text="Remove Task", command=self.remove_task)
        self.remove_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.update_button = ttk.Button(self.frame, text="Update Task", command=self.update_task)
        self.update_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.export_new_button = ttk.Button(self.frame, text="Export to New Excel", command=self.export_new_schedule)
        self.export_new_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.export_existing_button = ttk.Button(self.frame, text="Export to Existing Excel", command=self.export_existing_schedule)
        self.export_existing_button.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="we")

    # Function to add a task to the schedule
    def add_task(self):
        # Get task and time from the entry fields
        task = self.task_entry.get()
        time = self.time_entry.get()

        # Check if both task and time are provided
        if task and time:
            # Add the task to the schedule dictionary
            self.schedule[time] = task
            messagebox.showinfo("Success", "Task added successfully!")
            # Clear entry fields after adding task
            self.clear_entries()
        else:
            messagebox.showerror("Error", "Both task and time are required.")

    # Function to display the schedule from an Excel file
    # Function to display the schedule from the specified Excel file
    # Function to display the schedule from the specified Excel file
    def display_schedule(self):
        excel_file_path = r"C:\Users\Acer\OneDrive\Desktop\New folder\aryaschedule.xlsx"
        try:
            # Read the schedule from the Excel file into a DataFrame
            schedule_df = pd.read_excel(excel_file_path)
            # Check if the DataFrame is empty
            if not schedule_df.empty:
                # Construct schedule text for display
                schedule_text = "Your Schedule for Today:\n"
                for index, row in schedule_df.iterrows():
                    schedule_text += f"{row['Time']}: {row['Task']}\n"
                messagebox.showinfo("Schedule", schedule_text)
            else:
                messagebox.showinfo("Schedule", "No tasks scheduled for today.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while reading the schedule from Excel: {e}")

    # Function to remove a task from the schedule
    def remove_task(self):
        # Get time from the entry field
        time = self.time_entry.get()

        # Check if the task exists for the provided time
        if time in self.schedule:
            # Remove the task from the schedule
            del self.schedule[time]
            messagebox.showinfo("Success", "Task removed successfully!")
            # Clear entry fields after removing task
            self.clear_entries()
        else:
            messagebox.showerror("Error", "No task found at that time.")

    # Function to update a task in the schedule
    def update_task(self):
        # Get time and new task from the entry fields
        time = self.time_entry.get()
        new_task = self.task_entry.get()

        # Check if the task exists for the provided time
        if time in self.schedule:
            # Update the task in the schedule
            self.schedule[time] = new_task
            messagebox.showinfo("Success", "Task updated successfully!")
            # Clear entry fields after updating task
            self.clear_entries()
        else:
            messagebox.showerror("Error", "No task found at that time.")

    # Function to export the schedule to a new Excel file
    def export_new_schedule(self):
        self.export_schedule()

    # Function to export the schedule to an existing Excel file
    def export_existing_schedule(self):
        if self.schedule:
            # Ask user to select an existing Excel file
            file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                try:
                    # Check if the selected file exists
                    if os.path.exists(file_path):
                        # Read the existing Excel file into a DataFrame
                        existing_df = pd.read_excel(file_path)
                        # Create a DataFrame for the new entries
                        new_df = pd.DataFrame(list(self.schedule.items()), columns=['Time', 'Task'])
                        # Concatenate existing and new DataFrames and save to Excel
                        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
                        updated_df.to_excel(file_path, index=False)
                        messagebox.showinfo("Export Successful", f"Schedule appended to {file_path} successfully!")
                    else:
                        messagebox.showerror("Export Error", "Selected file does not exist.")
                except Exception as e:
                    messagebox.showerror("Export Error", f"An error occurred while exporting the schedule: {e}")
        else:
            messagebox.showwarning("Export Warning", "No schedule to export.")

    # Function to clear entry fields
    def clear_entries(self):
        self.task_entry.delete(0, tk.END)
        self.time_entry.delete(0, tk.END)

# Main function to create and run the application
def main():
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
