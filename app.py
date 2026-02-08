import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from scheduler import schedule_month
from export_excel import export_for_powerbi
from storage import load_json, save_json
from datetime import datetime
from models import Employee, PairingRecord, ISOAction
import logging
import sys
import os
import json


class LPAApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # Setup style
        self.setup_style()
        #For configuring the main window
        self.configure(bg='#F2F2F2')
        self.title("Digital LPA & ISO Compliance System")
        self.geometry("1200x700")

        # Setup logging
        self.setup_logging()



        # Initialize data
        self.employees = []
        self.sections = []
        self.shifts = [1, 2, 3]
        self.iso_actions = []
        self.pairing_history = []

        # Load saved data
        self.load_data()

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=5)

        self.setup_tab()
        self.scheduler_tab()
        self.iso_tab()
        self.add_utility_buttons()

        logging.info("Application started successfully")

    def setup_logging(self):
        log_path = os.path.join(os.getcwd(), "app_debug.log")
        logging.basicConfig(
            filename=log_path,
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def setup_style(self):
        """Apply Hitachi Energy color scheme"""
        style = ttk.Style()

        # Hitachi Energy Colors (from their brand guidelines)
        primary_blue = "#0D4D90"  # Primary blue
        accent_red = "#E4002B"  # Accent red
        light_gray = "#F2F2F2"  # Light background
        dark_gray = "#333333"  # Dark text
        white = "#FFFFFF"  # White

        # Configure overall theme
        style.theme_use('clam')  # 'clam' theme is most customizable

        # Configure colors for all ttk widgets
        style.configure('.',
                        background=light_gray,
                        foreground=dark_gray,
                        fieldbackground=white,
                        font=('Segoe UI', 10))

        # Configure specific widgets
        style.configure('TNotebook', background=primary_blue)
        style.configure('TNotebook.Tab',
                        background="#E6E6E6",
                        foreground=dark_gray,
                        padding=[15, 5])
        style.map('TNotebook.Tab',
                  background=[('selected', primary_blue)],
                  foreground=[('selected', white)])

        # Configure buttons
        style.configure('TButton',
                        background=primary_blue,
                        foreground=white,
                        borderwidth=1,
                        relief="raised",
                        padding=6)
        style.map('TButton',
                  background=[('active', accent_red), ('pressed', accent_red)])

        # Configure labels
        style.configure('TLabel',
                        background=light_gray,
                        foreground=dark_gray,
                        font=('Segoe UI', 10, 'bold'))

        # Configure entry fields
        style.configure('TEntry',
                        fieldbackground=white,
                        foreground=dark_gray,
                        insertcolor=primary_blue)

        # Configure combobox
        style.configure('TCombobox',
                        fieldbackground=white,
                        background=white,
                        arrowcolor=primary_blue)

        # Configure treeview
        style.configure('Treeview',
                        background=white,
                        foreground=dark_gray,
                        fieldbackground=white,
                        rowheight=25)
        style.configure('Treeview.Heading',
                        background=primary_blue,
                        foreground=white,
                        font=('Segoe UI', 10, 'bold'))
        style.map('Treeview',
                  background=[('selected', accent_red)],
                  foreground=[('selected', white)])

        # Configure scrollbars
        style.configure('Vertical.TScrollbar',
                        background=light_gray,
                        troughcolor=light_gray,
                        arrowcolor=primary_blue)
        style.configure('Horizontal.TScrollbar',
                        background=light_gray,
                        troughcolor=light_gray,
                        arrowcolor=primary_blue)

        # Configure frames
        style.configure('TFrame', background=light_gray)
        style.configure('TLabelframe', background=light_gray)
        style.configure('TLabelframe.Label',
                        background=light_gray,
                        foreground=primary_blue,
                        font=('Segoe UI', 10, 'bold'))

    def style_dialog(self, dialog):
        """Apply consistent styling to dialogs"""
        dialog.configure(bg='#F2F2F2')

        # Style all ttk widgets in the dialog
        for child in dialog.winfo_children():
            if isinstance(child, ttk.Frame):
                child.configure(style='TFrame')

    def load_data(self):
        """Load saved data from JSON files"""
        try:
            # Load employees
            employees_data = load_json("employees.json", [])
            self.employees = [
                Employee(**emp) for emp in employees_data
            ]

            # Load sections
            self.sections = load_json("sections.json", ["311", "341", "361"])

            # Load ISO actions
            iso_data = load_json("iso_actions.json", [])
            self.iso_actions = [
                ISOAction(**action) for action in iso_data
            ]

            # Load pairing history
            history_data = load_json("pairing_history.json", [])
            self.pairing_history = [
                PairingRecord(**record) for record in history_data
            ]

            logging.info(
                f"Loaded {len(self.employees)} employees, {len(self.sections)} sections, {len(self.iso_actions)} ISO actions, {len(self.pairing_history)} history records")

        except Exception as e:
            logging.error(f"Error loading data: {str(e)}")
            # Start with default employees if none loaded
            if not self.employees:
                self.employees = [
                    Employee("John", "Smith", "Engineer", 1),
                    Employee("Maria", "Lopez", "Supervisor", 2),
                    Employee("David", "Chen", "Manager", 3),
                ]

    def save_data(self):
        """Save current data to JSON files"""
        try:
            # Save employees
            employees_data = [
                {
                    "first_name": e.first_name,
                    "last_name": e.last_name,
                    "role": e.role,
                    "shift": e.shift,
                    "active": e.active
                }
                for e in self.employees
            ]
            save_json("employees.json", employees_data)

            # Save sections
            save_json("sections.json", self.sections)

            # Save ISO actions
            iso_data = [
                {
                    "description": a.description,
                    "owner": a.owner,
                    "due_date": a.due_date,
                    "status": a.status
                }
                for a in self.iso_actions
            ]
            save_json("iso_actions.json", iso_data)

            # Save pairing history
            history_data = [
                {
                    "emp_a": r.emp_a,
                    "emp_b": r.emp_b,
                    "month": r.month,
                    "year": r.year
                }
                for r in self.pairing_history
            ]
            save_json("pairing_history.json", history_data)

            logging.info("Data saved successfully")
            messagebox.showinfo("Saved", "Data has been saved successfully.")

        except Exception as e:
            logging.error(f"Error saving data: {str(e)}")
            messagebox.showerror("Save Error", f"Could not save data: {str(e)}")

    def setup_tab(self):
        """Setup tab for managing employees and sections"""
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Setup")

        # Create notebook within setup tab
        setup_notebook = ttk.Notebook(tab)
        setup_notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Employees frame
        employees_frame = ttk.Frame(setup_notebook)
        setup_notebook.add(employees_frame, text="Employees")
        self.setup_employees_tab(employees_frame)

        # Sections frame
        sections_frame = ttk.Frame(setup_notebook)
        setup_notebook.add(sections_frame, text="Sections")
        self.setup_sections_tab(sections_frame)

        # Save button
        ttk.Button(tab, text="💾 Save All Data", command=self.save_data).pack(pady=10)

    def setup_employees_tab(self, parent):
        """Create employee management interface"""
        # Treeview for employees
        columns = ("ID", "First Name", "Last Name", "Role", "Shift", "Active")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=15)

        # Define headings
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        tree.column("ID", width=50)
        tree.column("Active", width=60)

        # Scrollbar
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Button frame
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)

        # Buttons
        ttk.Button(button_frame, text="➕ Add Employee",
                   command=lambda: self.add_employee_dialog(tree)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="✏️ Edit Employee",
                   command=lambda: self.edit_employee_dialog(tree)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="🗑️ Delete Employee",
                   command=lambda: self.delete_employee(tree)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="🔄 Toggle Active",
                   command=lambda: self.toggle_employee_active(tree)).pack(side="left", padx=5)

        # Load employees into tree
        self.refresh_employees_tree(tree)

        # Configure grid weights
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

    def refresh_employees_tree(self, tree):
        """Refresh the employees treeview"""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)

        # Add employees
        for idx, emp in enumerate(self.employees, 1):
            tree.insert("", "end", values=(
                idx,
                emp.first_name,
                emp.last_name,
                emp.role,
                emp.shift,
                "✓" if emp.active else "✗"
            ))

    def add_employee_dialog(self, tree):
        """Dialog to add a new employee"""
        dialog = tk.Toplevel(self)
        dialog.title("Add New Employee")
        dialog.geometry("400x400")
        dialog.transient(self)
        dialog.grab_set()

        #For styling
        self.style_dialog(dialog)

        # Form fields
        ttk.Label(dialog, text="First Name:").pack(pady=(20, 5))
        first_name_entry = ttk.Entry(dialog, width=30)
        first_name_entry.pack(pady=5)

        ttk.Label(dialog, text="Last Name:").pack(pady=(10, 5))
        last_name_entry = ttk.Entry(dialog, width=30)
        last_name_entry.pack(pady=5)

        ttk.Label(dialog, text="Role:").pack(pady=(10, 5))
        role_entry = ttk.Entry(dialog, width=30)
        role_entry.pack(pady=5)

        ttk.Label(dialog, text="Shift (1, 2, or 3):").pack(pady=(10, 5))
        shift_entry = ttk.Entry(dialog, width=30)
        shift_entry.pack(pady=5)

        active_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(dialog, text="Active", variable=active_var).pack(pady=10)

        def save_employee():
            try:
                first_name = first_name_entry.get().strip()
                last_name = last_name_entry.get().strip()
                role = role_entry.get().strip()
                shift = int(shift_entry.get())

                if not first_name or not last_name or not role:
                    messagebox.showwarning("Validation", "Please fill in all fields")
                    return

                if shift not in [1, 2, 3]:
                    messagebox.showwarning("Validation", "Shift must be 1, 2, or 3")
                    return

                new_employee = Employee(
                    first_name=first_name,
                    last_name=last_name,
                    role=role,
                    shift=shift,
                    active=active_var.get()
                )

                self.employees.append(new_employee)
                self.refresh_employees_tree(tree)
                logging.info(f"Added employee: {new_employee.name()}")

                dialog.destroy()
                messagebox.showinfo("Success", "Employee added successfully!")

            except ValueError:
                messagebox.showerror("Error", "Shift must be a number (1, 2, or 3)")
            except Exception as e:
                messagebox.showerror("Error", f"Could not add employee: {str(e)}")

        ttk.Button(dialog, text="Save", command=save_employee).pack(pady=20)

    def edit_employee_dialog(self, tree):
        """Dialog to edit selected employee"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an employee to edit")
            return

        item = tree.item(selection[0])
        values = item['values']
        emp_index = values[0] - 1  # Convert display ID to list index

        if emp_index >= len(self.employees):
            messagebox.showerror("Error", "Invalid employee selection")
            return

        employee = self.employees[emp_index]

        dialog = tk.Toplevel(self)
        dialog.title("Edit Employee")
        dialog.geometry("400x300")
        dialog.transient(self)
        dialog.grab_set()

        # Form fields with current values
        ttk.Label(dialog, text="First Name:").pack(pady=(20, 5))
        first_name_entry = ttk.Entry(dialog, width=30)
        first_name_entry.insert(0, employee.first_name)
        first_name_entry.pack(pady=5)

        ttk.Label(dialog, text="Last Name:").pack(pady=(10, 5))
        last_name_entry = ttk.Entry(dialog, width=30)
        last_name_entry.insert(0, employee.last_name)
        last_name_entry.pack(pady=5)

        ttk.Label(dialog, text="Role:").pack(pady=(10, 5))
        role_entry = ttk.Entry(dialog, width=30)
        role_entry.insert(0, employee.role)
        role_entry.pack(pady=5)

        ttk.Label(dialog, text="Shift (1, 2, or 3):").pack(pady=(10, 5))
        shift_entry = ttk.Entry(dialog, width=30)
        shift_entry.insert(0, str(employee.shift))
        shift_entry.pack(pady=5)

        active_var = tk.BooleanVar(value=employee.active)
        ttk.Checkbutton(dialog, text="Active", variable=active_var).pack(pady=10)

        def save_changes():
            try:
                employee.first_name = first_name_entry.get().strip()
                employee.last_name = last_name_entry.get().strip()
                employee.role = role_entry.get().strip()
                employee.shift = int(shift_entry.get())
                employee.active = active_var.get()

                if not employee.first_name or not employee.last_name or not employee.role:
                    messagebox.showwarning("Validation", "Please fill in all fields")
                    return

                if employee.shift not in [1, 2, 3]:
                    messagebox.showwarning("Validation", "Shift must be 1, 2, or 3")
                    return

                self.refresh_employees_tree(tree)
                logging.info(f"Edited employee: {employee.name()}")

                dialog.destroy()
                messagebox.showinfo("Success", "Employee updated successfully!")

            except ValueError:
                messagebox.showerror("Error", "Shift must be a number (1, 2, or 3)")
            except Exception as e:
                messagebox.showerror("Error", f"Could not update employee: {str(e)}")

        ttk.Button(dialog, text="Save Changes", command=save_changes).pack(pady=20)

    def delete_employee(self, tree):
        """Delete selected employee"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an employee to delete")
            return

        item = tree.item(selection[0])
        values = item['values']
        emp_index = values[0] - 1

        if emp_index >= len(self.employees):
            messagebox.showerror("Error", "Invalid employee selection")
            return

        employee = self.employees[emp_index]

        response = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete {employee.name()}?"
        )

        if response:
            del self.employees[emp_index]
            self.refresh_employees_tree(tree)
            logging.info(f"Deleted employee: {employee.name()}")
            messagebox.showinfo("Deleted", "Employee deleted successfully!")

    def toggle_employee_active(self, tree):
        """Toggle active status of selected employee"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an employee")
            return

        item = tree.item(selection[0])
        values = item['values']
        emp_index = values[0] - 1

        if emp_index >= len(self.employees):
            messagebox.showerror("Error", "Invalid employee selection")
            return

        employee = self.employees[emp_index]
        employee.active = not employee.active

        self.refresh_employees_tree(tree)
        status = "activated" if employee.active else "deactivated"
        logging.info(f"{status} employee: {employee.name()}")
        messagebox.showinfo("Status Changed", f"Employee {status} successfully!")

    def setup_sections_tab(self, parent):
        """Create sections management interface"""
        # Listbox for sections
        listbox_frame = ttk.Frame(parent)
        listbox_frame.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(listbox_frame, text="Plant Sections:", font=("Arial", 10, "bold")).pack(anchor="w")

        # Listbox with scrollbar
        listbox = tk.Listbox(listbox_frame, height=10, font=("Arial", 10))
        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)

        listbox.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")

        # Button frame
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=10)

        # Buttons
        ttk.Button(button_frame, text="➕ Add Section",
                   command=lambda: self.add_section_dialog(listbox)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="🗑️ Delete Section",
                   command=lambda: self.delete_section(listbox)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="✏️ Edit Section",
                   command=lambda: self.edit_section_dialog(listbox)).pack(side="left", padx=5)

        # Load sections into listbox
        self.refresh_sections_listbox(listbox)

    def refresh_sections_listbox(self, listbox):
        """Refresh the sections listbox"""
        listbox.delete(0, tk.END)
        for section in self.sections:
            listbox.insert(tk.END, section)

    def add_section_dialog(self, listbox):
        """Dialog to add a new section"""
        section = simpledialog.askstring("Add Section", "Enter new section code:")
        if section and section.strip():
            section = section.strip()
            if section not in self.sections:
                self.sections.append(section)
                self.refresh_sections_listbox(listbox)
                logging.info(f"Added section: {section}")
                messagebox.showinfo("Success", f"Section '{section}' added successfully!")
            else:
                messagebox.showwarning("Duplicate", "Section already exists!")

    def delete_section(self, listbox):
        """Delete selected section"""
        selection = listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a section to delete")
            return

        section = listbox.get(selection[0])
        response = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete section '{section}'?"
        )

        if response:
            self.sections.remove(section)
            self.refresh_sections_listbox(listbox)
            logging.info(f"Deleted section: {section}")
            messagebox.showinfo("Deleted", f"Section '{section}' deleted successfully!")

    def edit_section_dialog(self, listbox):
        """Dialog to edit selected section"""
        selection = listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a section to edit")
            return

        old_section = listbox.get(selection[0])
        new_section = simpledialog.askstring("Edit Section", "Enter new section code:", initialvalue=old_section)

        if new_section and new_section.strip():
            new_section = new_section.strip()
            if new_section == old_section:
                return  # No change

            if new_section in self.sections:
                messagebox.showwarning("Duplicate", "Section already exists!")
                return

            index = self.sections.index(old_section)
            self.sections[index] = new_section
            self.refresh_sections_listbox(listbox)
            logging.info(f"Edited section: {old_section} -> {new_section}")
            messagebox.showinfo("Success", f"Section updated to '{new_section}'!")

    def scheduler_tab(self):
        """Enhanced scheduler tab with month selection and results display"""
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="LPA Scheduler")

        # Control frame for month selection and buttons
        control_frame = ttk.Frame(tab)
        control_frame.pack(fill="x", padx=20, pady=10)

        # Month selection
        ttk.Label(control_frame, text="Month:").grid(row=0, column=0, padx=(0, 5))

        month_var = tk.StringVar()
        month_combo = ttk.Combobox(control_frame, textvariable=month_var,
                                   values=list(range(1, 13)), width=5, state="readonly")
        month_combo.set(str(datetime.now().month))
        month_combo.grid(row=0, column=1, padx=(0, 20))

        # Year selection
        ttk.Label(control_frame, text="Year:").grid(row=0, column=2, padx=(0, 5))

        year_var = tk.StringVar()
        current_year = datetime.now().year
        year_combo = ttk.Combobox(control_frame, textvariable=year_var,
                                  values=list(range(current_year - 1, current_year + 3)),
                                  width=6, state="readonly")
        year_combo.set(current_year)
        year_combo.grid(row=0, column=3, padx=(0, 20))

        # Results frame
        results_frame = ttk.LabelFrame(tab, text="Schedule Results", padding=10)
        results_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Treeview for results
        results_columns = ("Section", "Target Shift", "Auditors", "Auditor Count")
        results_tree = ttk.Treeview(results_frame, columns=results_columns, show="headings", height=12)

        for col in results_columns:
            results_tree.heading(col, text=col)
            results_tree.column(col, width=150)

        # Scrollbars
        y_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=results_tree.yview)
        x_scrollbar = ttk.Scrollbar(results_frame, orient="horizontal", command=results_tree.xview)
        results_tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

        results_tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar.grid(row=1, column=0, sticky="ew", columnspan=2)

        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        def run_scheduler():
            try:
                if not self.employees:
                    messagebox.showwarning("No Employees", "Please add employees in the Setup tab first.")
                    return

                if not self.sections:
                    messagebox.showwarning("No Sections", "Please add sections in the Setup tab first.")
                    return

                month = int(month_var.get())
                year = int(year_var.get())

                logging.info(f"Creating schedule for {month}/{year}")

                # Get active employees
                active_employees = [e for e in self.employees if e.active]

                if len(active_employees) < 2:
                    messagebox.showwarning("Not Enough Auditors",
                                           "Need at least 2 active employees to generate schedule.")
                    return

                # Run scheduler
                result = schedule_month(
                    active_employees,
                    self.sections,
                    self.shifts,
                    self.pairing_history,
                    month,
                    year
                )

                # Update pairing history with new pairings
                for assignment in result:
                    if len(assignment.auditors) == 2:
                        self.pairing_history.append(PairingRecord(
                            assignment.auditors[0],
                            assignment.auditors[1],
                            month,
                            year
                        ))

                # Display results
                self.display_schedule_results(results_tree, result)

                # Export to Excel
                iso_data = [
                    {
                        "description": a.description,
                        "owner": a.owner,
                        "due_date": a.due_date,
                        "status": a.status,
                        "overdue": False  # You can add overdue logic here
                    }
                    for a in self.iso_actions
                ]

                filepath = export_for_powerbi(result, iso_data)

                if filepath:
                    messagebox.showinfo(
                        "Success",
                        f"Schedule created successfully!\n\n"
                        f"Total assignments: {len(result)}\n"
                        f"File saved to:\n{filepath}"
                    )
                    logging.info("Export completed successfully")

            except Exception as e:
                error_msg = f"Error generating schedule: {str(e)}"
                logging.error(error_msg)
                messagebox.showerror("Error", error_msg)

        def display_summary():
            """Display summary of employees and sections"""
            active_count = sum(1 for e in self.employees if e.active)
            summary = (
                f"Active Employees: {active_count}/{len(self.employees)}\n"
                f"Sections: {len(self.sections)}\n"
                f"ISO Actions: {len(self.iso_actions)}\n"
                f"Pairing History Records: {len(self.pairing_history)}"
            )
            messagebox.showinfo("System Summary", summary)

        # Buttons
        ttk.Button(control_frame, text="📊 Generate Schedule",
                   command=run_scheduler).grid(row=0, column=4, padx=(20, 5))
        ttk.Button(control_frame, text="📈 View Summary",
                   command=display_summary).grid(row=0, column=5, padx=5)
        ttk.Button(control_frame, text="🔄 Clear Results",
                   command=lambda: self.clear_results_tree(results_tree)).grid(row=0, column=6, padx=5)

    def display_schedule_results(self, tree, assignments):
        """Display schedule results in treeview"""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)

        # Add assignments
        for assignment in assignments:
            tree.insert("", "end", values=(
                assignment.section,
                assignment.target_shift,
                ", ".join(assignment.auditors),
                len(assignment.auditors)
            ))

    def clear_results_tree(self, tree):
        """Clear the results treeview"""
        for item in tree.get_children():
            tree.delete(item)

    def iso_tab(self):
        """ISO/Corrective Actions management tab"""
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="ISO / Corrective Actions")

        # Treeview for ISO actions
        columns = ("Description", "Owner", "Due Date", "Status")
        tree = ttk.Treeview(tab, columns=columns, show="headings", height=15)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=200)

        tree.column("Owner", width=150)
        tree.column("Due Date", width=100)
        tree.column("Status", width=80)

        # Scrollbars
        y_scrollbar = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
        x_scrollbar = ttk.Scrollbar(tab, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

        tree.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar.grid(row=1, column=0, sticky="ew", columnspan=2)

        # Button frame
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        # Buttons
        ttk.Button(button_frame, text="➕ Add Action",
                   command=lambda: self.add_iso_action_dialog(tree)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="✏️ Edit Action",
                   command=lambda: self.edit_iso_action_dialog(tree)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="🗑️ Delete Action",
                   command=lambda: self.delete_iso_action(tree)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="✅ Mark Complete",
                   command=lambda: self.mark_iso_action_complete(tree)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="🔄 Refresh",
                   command=lambda: self.refresh_iso_tree(tree)).pack(side="left", padx=5)

        # Configure grid weights
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        # Load initial data
        self.refresh_iso_tree(tree)

    def refresh_iso_tree(self, tree):
        """Refresh the ISO actions treeview"""
        for item in tree.get_children():
            tree.delete(item)

        for action in self.iso_actions:
            # Color code overdue actions
            tree.insert("", "end", values=(
                action.description,
                action.owner,
                action.due_date,
                action.status
            ))

    def add_iso_action_dialog(self, tree):
        """Dialog to add new ISO action"""
        dialog = tk.Toplevel(self)
        dialog.title("Add ISO Action")
        dialog.geometry("500x400")
        dialog.transient(self)
        dialog.grab_set()

        # Form fields
        ttk.Label(dialog, text="Description:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(20, 5))
        desc_text = tk.Text(dialog, height=4, width=50)
        desc_text.pack(padx=20, pady=5)

        ttk.Label(dialog, text="Owner:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(10, 5))
        owner_entry = ttk.Entry(dialog, width=50)
        owner_entry.pack(padx=20, pady=5)

        ttk.Label(dialog, text="Due Date (YYYY-MM-DD):", font=("Arial", 10, "bold")).pack(anchor="w", padx=20,
                                                                                          pady=(10, 5))
        due_entry = ttk.Entry(dialog, width=50)
        due_entry.pack(padx=20, pady=5)

        ttk.Label(dialog, text="Initial Status:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(10, 5))
        status_var = tk.StringVar(value="Open")
        status_combo = ttk.Combobox(dialog, textvariable=status_var,
                                    values=["Open", "Closed", "In Progress"],
                                    state="readonly", width=47)
        status_combo.pack(padx=20, pady=5)

        def save_action():
            try:
                description = desc_text.get("1.0", tk.END).strip()
                owner = owner_entry.get().strip()
                due_date = due_entry.get().strip()
                status = status_var.get()

                if not description or not owner or not due_date:
                    messagebox.showwarning("Validation", "Please fill in all fields")
                    return

                new_action = ISOAction(
                    description=description,
                    owner=owner,
                    due_date=due_date,
                    status=status
                )

                self.iso_actions.append(new_action)
                self.refresh_iso_tree(tree)
                logging.info(f"Added ISO action: {description[:50]}...")

                dialog.destroy()
                messagebox.showinfo("Success", "ISO action added successfully!")

            except Exception as e:
                messagebox.showerror("Error", f"Could not add action: {str(e)}")

        ttk.Button(dialog, text="Save Action", command=save_action).pack(pady=20)

    def edit_iso_action_dialog(self, tree):
        """Dialog to edit selected ISO action"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an action to edit")
            return

        item = tree.item(selection[0])
        values = item['values']

        # Find the action
        action = None
        for a in self.iso_actions:
            if (a.description == values[0] and a.owner == values[1] and
                    a.due_date == values[2] and a.status == values[3]):
                action = a
                break

        if not action:
            messagebox.showerror("Error", "Action not found")
            return

        dialog = tk.Toplevel(self)
        dialog.title("Edit ISO Action")
        dialog.geometry("500x400")
        dialog.transient(self)
        dialog.grab_set()

        # Form fields with current values
        ttk.Label(dialog, text="Description:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(20, 5))
        desc_text = tk.Text(dialog, height=4, width=50)
        desc_text.insert("1.0", action.description)
        desc_text.pack(padx=20, pady=5)

        ttk.Label(dialog, text="Owner:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(10, 5))
        owner_entry = ttk.Entry(dialog, width=50)
        owner_entry.insert(0, action.owner)
        owner_entry.pack(padx=20, pady=5)

        ttk.Label(dialog, text="Due Date (YYYY-MM-DD):", font=("Arial", 10, "bold")).pack(anchor="w", padx=20,
                                                                                          pady=(10, 5))
        due_entry = ttk.Entry(dialog, width=50)
        due_entry.insert(0, action.due_date)
        due_entry.pack(padx=20, pady=5)

        ttk.Label(dialog, text="Status:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(10, 5))
        status_var = tk.StringVar(value=action.status)
        status_combo = ttk.Combobox(dialog, textvariable=status_var,
                                    values=["Open", "Closed", "In Progress"],
                                    state="readonly", width=47)
        status_combo.pack(padx=20, pady=5)

        def save_changes():
            try:
                action.description = desc_text.get("1.0", tk.END).strip()
                action.owner = owner_entry.get().strip()
                action.due_date = due_entry.get().strip()
                action.status = status_var.get()

                if not action.description or not action.owner or not action.due_date:
                    messagebox.showwarning("Validation", "Please fill in all fields")
                    return

                self.refresh_iso_tree(tree)
                logging.info(f"Edited ISO action: {action.description[:50]}...")

                dialog.destroy()
                messagebox.showinfo("Success", "ISO action updated successfully!")

            except Exception as e:
                messagebox.showerror("Error", f"Could not update action: {str(e)}")

        ttk.Button(dialog, text="Save Changes", command=save_changes).pack(pady=20)

    def delete_iso_action(self, tree):
        """Delete selected ISO action"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an action to delete")
            return

        item = tree.item(selection[0])
        values = item['values']

        # Find and remove the action
        for i, action in enumerate(self.iso_actions):
            if (action.description == values[0] and action.owner == values[1] and
                    action.due_date == values[2] and action.status == values[3]):

                response = messagebox.askyesno(
                    "Confirm Delete",
                    f"Are you sure you want to delete this action?\n\n"
                    f"Description: {action.description[:50]}..."
                )

                if response:
                    del self.iso_actions[i]
                    self.refresh_iso_tree(tree)
                    logging.info(f"Deleted ISO action: {action.description[:50]}...")
                    messagebox.showinfo("Deleted", "Action deleted successfully!")
                return

        messagebox.showerror("Error", "Action not found")

    def mark_iso_action_complete(self, tree):
        """Mark selected ISO action as complete"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an action to mark complete")
            return

        item = tree.item(selection[0])
        values = item['values']

        # Find the action
        for action in self.iso_actions:
            if (action.description == values[0] and action.owner == values[1] and
                    action.due_date == values[2] and action.status == values[3]):

                if action.status == "Closed":
                    messagebox.showinfo("Already Closed", "This action is already closed.")
                    return

                action.status = "Closed"
                self.refresh_iso_tree(tree)
                logging.info(f"Marked ISO action as complete: {action.description[:50]}...")
                messagebox.showinfo("Success", "Action marked as complete!")
                return

        messagebox.showerror("Error", "Action not found")

    def add_utility_buttons(self):
        tab = ttk.Frame(self.tabs)
        self.tabs.add(tab, text="Utilities")

        def view_log():
            try:
                log_path = os.path.join(os.getcwd(), "app_debug.log")
                if os.path.exists(log_path):
                    import subprocess
                    subprocess.Popen(['notepad.exe', log_path])
                else:
                    messagebox.showinfo("Info", "No log file found yet.")
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open log: {str(e)}")

        def open_output_folder():
            try:
                if getattr(sys, 'frozen', False):
                    folder = os.path.join(os.path.dirname(sys.executable), "outputs")
                else:
                    folder = os.path.join(os.getcwd(), "outputs")

                os.makedirs(folder, exist_ok=True)
                import subprocess
                subprocess.Popen(f'explorer "{folder}"')
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open folder: {str(e)}")

        def backup_data():
            """Create a backup of all data"""
            try:
                import shutil
                import time

                timestamp = time.strftime("%Y%m%d_%H%M%S")
                backup_dir = os.path.join(os.getcwd(), "backups")
                os.makedirs(backup_dir, exist_ok=True)

                data_files = ["employees.json", "sections.json", "iso_actions.json", "pairing_history.json"]

                for file in data_files:
                    src = os.path.join("data", file)
                    if os.path.exists(src):
                        dst = os.path.join(backup_dir, f"{timestamp}_{file}")
                        shutil.copy2(src, dst)

                logging.info(f"Created backup: {timestamp}")
                messagebox.showinfo("Backup", f"Backup created successfully!\nTimestamp: {timestamp}")

            except Exception as e:
                logging.error(f"Backup error: {str(e)}")
                messagebox.showerror("Backup Error", f"Could not create backup: {str(e)}")

        ttk.Button(tab, text="📋 View Log File", command=view_log).pack(pady=10)
        ttk.Button(tab, text="📂 Open Output Folder", command=open_output_folder).pack(pady=10)
        ttk.Button(tab, text="💾 Backup All Data", command=backup_data).pack(pady=10)
        ttk.Button(tab, text="🗑️ Clear Log File",
                   command=lambda: self.clear_log_file()).pack(pady=10)

    def clear_log_file(self):
        """Clear the log file"""
        response = messagebox.askyesno("Clear Log", "Are you sure you want to clear the log file?")
        if response:
            try:
                log_path = os.path.join(os.getcwd(), "app_debug.log")
                with open(log_path, 'w') as f:
                    f.write("")
                messagebox.showinfo("Log Cleared", "Log file has been cleared.")
                logging.info("Log file cleared by user")
            except Exception as e:
                messagebox.showerror("Error", f"Could not clear log: {str(e)}")



if __name__ == "__main__":
    app = LPAApp()
    app.mainloop()