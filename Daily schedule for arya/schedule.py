import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog 
import pandas as pd

class ScheduleApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Daily Schedule App")
        
        self.schedule = {}  # Initialize an empty dictionary to store schedule data
        
        self.create_widgets()  # Create GUI widgets

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

        # Buttons for adding, displaying, removing, and exporting schedule
        self.add_button = ttk.Button(self.frame, text="Add Task", command=self.add_task)
        self.add_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.display_button = ttk.Button(self.frame, text="Display Schedule", command=self.display_schedule)
        self.display_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.remove_button = ttk.Button(self.frame, text="Remove Task", command=self.remove_task)
        self.remove_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.export_button = ttk.Button(self.frame, text="Export to Excel", command=self.export_schedule)
        self.export_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="we")

    # Function to add a task to the schedule
    def add_task(self):
        task = self.task_entry.get()
        time = self.time_entry.get()

        if task and time:
            self.schedule[time] = task
            messagebox.showinfo("Success", "Task added successfully!")
            self.clear_entries()  # Clear entry fields after adding task
        else:
            messagebox.showerror("Error", "Both task and time are required.")

    # Function to display the schedule
    def display_schedule(self):
        if self.schedule:
            schedule_text = "Your Schedule for Today:\n"
            for time, task in self.schedule.items():
                schedule_text += f"{time}: {task}\n"
            messagebox.showinfo("Schedule", schedule_text)
        else:
            messagebox.showinfo("Schedule", "No tasks scheduled for today.")

    # Function to remove a task from the schedule
    def remove_task(self):
        time = self.time_entry.get()

        if time in self.schedule:
            del self.schedule[time]
            messagebox.showinfo("Success", "Task removed successfully!")
            self.clear_entries()  # Clear entry fields after removing task
        else:
            messagebox.showerror("Error", "No task found at that time.")

    # Function to export the schedule to an Excel file
    def export_schedule(self):
        if self.schedule:
            file_path = tk.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                try:
                    schedule_df = pd.DataFrame(list(self.schedule.items()), columns=['Time', 'Task'])
                    schedule_df.to_excel(file_path, index=False)
                    messagebox.showinfo("Export Successful", f"Schedule exported to \"add path here" successfully!")
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
