import tkinter as tk
from tkinter import StringVar, ttk, filedialog, messagebox
from csv import DictReader as csv_DictReader
from typing import List
from SendToPrint import DymoPrintService

class Employee():
    Employee = ""
    Guest_1 = ""
    Guest_2 = ""
    Guest_3 = ""
    Tour = ""
    ID=0

    # constructor
    def __init__(self, emp, g1, g2, g3, tour, id):
        self.Employee = emp
        self.Guest_1 = g1
        self.Guest_2 = g2
        self.Guest_3 = g3
        self.Tour = tour
        self.ID = id

    def updateDetails(self, emp, g1, g2, g3, tour):
        self.Employee = emp
        self.Guest_1 = g1
        self.Guest_2 = g2
        self.Guest_3 = g3
        self.Tour = tour
    
    def __getitem__(self, key):
        return {"Employee": self.Employee, "Guest_1": self.Guest_1, "Guest_2": self.Guest_2, "Guest_3": self.Guest_3, "Tour": self.Tour, "ID": self.ID}[key]


class DymoPrintManager(tk.Tk):
    printer = DymoPrintService()
    tour_options=["None"]
    
    def __init__(self):
        super().__init__()

        # Window dimensions
        w = 400
        h = 500

        # get screen width and height
        ws = self.winfo_screenwidth() # width of the screen
        hs = self.winfo_screenheight() # height of the screen

        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)

        # set the dimensions of the screen 
        # and where it is placed
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        # Set the title of the window
        self.title("Dymo Print Manager")

        #Configure grid for widget placement
        self.configure_grid()

        # Create and place the widgets
        self.setup_widgets()

        # List to store contacts
        self.employees: List[Employee]= []
        self.active_employees: List[Employee]= []

    def configure_grid(self):
        # Configure the rows and columns to stretch
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=1)
        self.rowconfigure(3, weight=1)

    def setup_widgets(self):
        # Search Field
        self.search_entry = ttk.Entry(self)
        self.search_entry.grid(column=0, row=0, columnspan=2, padx=10, pady=5, sticky='EW')

        # Search Button
        self.add_button = ttk.Button(self, text="Search", command=self.find_employee)
        self.add_button.grid(column=2, row=0, pady=5, padx=10, sticky='EW')

        # Filter By Tour Label
        self.filter_by_tour_label = ttk.Label(self, text="Filter By Tour:").grid(column=0, row=2, padx=10, pady=5, sticky='W')

        # Filter By Tour Menu
        self.selected_option = StringVar()
        self.selected_option.set(self.tour_options[0])
        self.filter_by_tour_menu = ttk.OptionMenu(self,self.selected_option,*self.tour_options)
        self.filter_by_tour_menu.grid(column=2, row=2, padx=10, pady=5, sticky='EW')
        
        self.file_button = ttk.Button(self, text="Load From CSV", command=self.choose_file)
        self.file_button.grid(column=0, row=4, padx=10, pady=10, sticky='W')

        self.print_button = ttk.Button(self, text="Print List", command=self.print_labels_from_listbox)
        self.print_button.grid(column=2, row=4, padx=10, pady=10, sticky='E')

        # Listbox to display contact names
        self.employee_listbox = tk.Listbox(self)
        self.employee_listbox.grid(column=0, row=3, columnspan=3, padx=10, pady=10, sticky="WESN")
        self.employee_listbox.bind("<<ListboxSelect>>", self.edit_contact_details)


    def choose_file(self):
        # Open file dialog to select a file
        file_path = filedialog.askopenfilename()

        # If a file is selected, update the label
        if file_path:
            postfix = file_path.split('.')[-1]
            if(postfix=='csv'):
                self.employees.clear() 
                self.employees = parse_csv(file_path)
                self.active_employees = self.employees
                self.update_tour_options()
            else:
                messagebox.showerror('Error', 'The file type you selected is not supported. Select a .csv file.')
            self.refresh_listbox()
        else:
            self.file_label.config(text="No file selected")

    ### Edit and Import Window Logic
    def edit_contact_details(self, event):
        # Get the selected contact index
        selection = event.widget.curselection()
        if not selection:
            return
        index = selection[0]
        contact = self.find_by_id(self.active_employees[index].ID)

        # Create a new Toplevel window for editing
        edit_window = tk.Toplevel(self)
        edit_window.title("Edit Employee Details")

        # Window dimensions
        w = 300
        h = 200

        # get screen width and height
        ws = self.winfo_screenwidth() # width of the screen
        hs = self.winfo_screenheight() # height of the screen

        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)

        edit_window.geometry('%dx%d+%d+%d' % (w, h, x, y))

        # Configure edit window grid
        edit_window.columnconfigure(0, weight=1)
        edit_window.columnconfigure(1, weight=1)
        edit_window.columnconfigure(2, weight=1)

        # Labels and Entry widgets for editing the contact details
        tk.Label(edit_window, text="Employee:").grid(column=0, row=0, padx=10, pady=5, sticky='W')
        employee_entry = ttk.Entry(edit_window)
        employee_entry.grid(column=1, row=0, columnspan=2, padx=10, pady=5, sticky='EW')
        employee_entry.insert(0, str(contact.Employee))

        tk.Label(edit_window, text="Guest 1:").grid(column=0, row=1, padx=10, pady=5, sticky='W')
        guest1_entry = ttk.Entry(edit_window)
        guest1_entry.grid(column=1, row=1, columnspan=2, padx=10, pady=5, sticky='EW')
        guest1_entry.insert(0, str(contact.Guest_1))

        tk.Label(edit_window, text="Guest 2:").grid(column=0, row=2, padx=10, pady=5, sticky='W')
        guest2_entry = ttk.Entry(edit_window)
        guest2_entry.grid(column=1, row=2, columnspan=2, padx=10, pady=5, sticky='EW')
        guest2_entry.insert(0, contact.Guest_2)

        tk.Label(edit_window, text="Guest 3:").grid(column=0, row=3, padx=10, pady=5, sticky='W')
        guest3_entry = ttk.Entry(edit_window)
        guest3_entry.grid(column=1, row=3, columnspan=2, padx=10, pady=5, sticky='EW')
        guest3_entry.insert(0, contact.Guest_3)

        tk.Label(edit_window, text="Tour:").grid(column=0, row=4, padx=10, pady=5, sticky='W')
        tour_entry = ttk.Entry(edit_window)
        tour_entry.grid(column=1, row=4, columnspan=2, padx=10, pady=5, sticky='EW')
        tour_entry.insert(0, contact.Tour)

        # Update Button
        update_button = ttk.Button(edit_window, text="Update", command=lambda: self.update_contact(contact.ID, employee_entry.get(), guest1_entry.get(), guest2_entry.get(), guest3_entry.get(), tour_entry.get(), edit_window))
        update_button.grid(column=0, row=5, pady=10)

        # Delete Button
        delete_button = ttk.Button(edit_window, text="Delete", command=lambda: self.delete_contact(contact.ID, edit_window))
        delete_button.grid(column=1, row=5, pady=10)

        # Print Button
        #delete_button = ttk.Button(edit_window, text="Print", command=lambda: self.printer.printLabelList([contact]))
        print_button = ttk.Button(edit_window, text="Print", command=lambda: self.print_contact(contact.ID, employee_entry.get(), guest1_entry.get(), guest2_entry.get(), guest3_entry.get(), tour_entry.get(), edit_window))
        print_button.grid(column=2, row=5, pady=10)

    ### Listbox Entry Handler Functions
    def update_contact(self, emp_id, employee, guest1, guest2, guest3, tour, edit_window):
        # Update the contact details in the contacts list
        emp_index = self.find_index_by_id(emp_id)
        self.employees[emp_index].updateDetails(employee, guest1, guest2, guest3, tour)

        # Refresh the Listbox and Tour dropdown to reflect the changes
        self.refresh_listbox(reset=False)
        self.update_tour_options()

        # Close the edit window
        edit_window.destroy()

    def delete_contact(self, emp_id, edit_window):
        # Delete the contact at the provided index
        emp_index = self.find_index_by_id(emp_id)
        self.employees.pop(emp_index)

        # Refresh the Listbox to reflect the changes
        self.refresh_listbox(reset=False)
        self.update_tour_options()

        # Close the edit window
        edit_window.destroy()
    
    def print_contact(self, emp_id, employee, guest1, guest2, guest3, tour, edit_window):
        emp_index = self.find_index_by_id(emp_id)
        selected_emp = self.find_by_id(emp_id)

        # If any changes have been made, call update before printing
        if selected_emp.Employee != employee or selected_emp.Guest_1 != guest1 or selected_emp.Guest_2 != guest2 or selected_emp.Guest_3 != guest3 or selected_emp.Tour != tour:
            self.update_contact(emp_index, employee, guest1, guest2, guest3, tour, edit_window)
        self.printer.printLabelList([self.employees[emp_index]])
    
    def refresh_listbox(self, reset=True):
        # Clear the Listbox
        self.employee_listbox.delete(0, tk.END)

        # Sync Lists
        # If not resetting the whole list, remove and re-add each active record to get updated info
        if not reset:
            results: List[Employee]=[]
            for e in self.active_employees:
                id = e.ID
                index = 0
                new_emp = self.find_by_id(id)
                if new_emp:
                    #print(new_emp.Employee)
                    results.append(new_emp)
                index += 1
            self.active_employees = results
        else:
            self.active_employees = self.employees

        # Sort Contacts
        self.active_employees.sort(key=lambda emp: emp.Employee)
        
        # Insert all contacts into the Listbox
        for emp in self.active_employees:
            self.employee_listbox.insert(tk.END, f"{emp.Employee}")

    def search_listbox(self, value, param):
        if param == "Name":
            self.selected_option.set("None")
            self.employee_listbox.delete(0, tk.END)
            self.active_employees = list(filter(lambda emp: str.lower(value) in str.lower(emp.Employee), self.employees))
            for emp in self.active_employees:
                self.employee_listbox.insert(tk.END, f"{emp.Employee}")
        elif param == "Tour":
            self.search_entry.delete(0,tk.END)
            self.employee_listbox.delete(0, tk.END)
            self.active_employees = list(filter(lambda emp: emp.Tour == value, self.employees))
            for emp in self.active_employees:
                self.employee_listbox.insert(tk.END, f"{emp.Employee}")

    def find_employee(self):
        self.search_listbox(self.search_entry.get(),"Name")

    def find_index_by_id(self,id):
        employee = list(filter(lambda emp: emp.ID == id, self.employees))[0]
        return self.employees.index(employee)
    
    def find_by_id(self,id):
        id = list(filter(lambda emp: emp.ID == id, self.employees))
        if len(id) != 0:
            return id[0]
        else:
            return False

    def update_selected_option(self,opt):
        self.selected_option.set(opt)
        if opt != "None":
            self.search_listbox(self.selected_option.get(),"Tour")
        else:
            self.refresh_listbox()

    def update_tour_options(self):
        tours=['None']
        self.tour_options = None
        for e in self.employees:
            if e.Tour not in tours:
                tours.append(e.Tour)
        self.tour_options = tours.sort()
        menu = self.filter_by_tour_menu["menu"]
        menu.delete(0, "end")
        for t in tours:
            menu.add_command(label=t,command=lambda value=t: self.update_selected_option(value))

    ### Print Handler Functions
    def print_labels_from_listbox(self):
        listbox_entries = self.employee_listbox.get(0,tk.END)
        entries_to_print=[]
        for e in listbox_entries:
            entries_to_print.append(next(filter(lambda emp: str.lower(e) in str.lower(emp.Employee), self.employees),None))
        self.printer.printLabelList(entries_to_print)

### Data Handler Functions
def parse_csv(file_path):
    employees = []
    index=0
    for row in csv_DictReader(open(file_path, encoding='utf-8-sig')):
        index+=1
        employees.append(Employee(f"{row['First_Name']} {row['Surename']}",row['Guest_1'],row['Guest_2'],row['Guest_3'],row['Tour_Number'],row['Employee_KOID']))
    
    return employees


if __name__ == "__main__":
    app = DymoPrintManager()
    app.mainloop()