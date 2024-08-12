import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl


formField_list = []
cwd = os.getcwd()
def submit_data():
    accepted = formField_list[9].get()
    
    if accepted=="Accepted":
        # User info
        firstname = formField_list[0].get()
        lastname = formField_list[1].get()
        balance = 0
        total_cost = 5000
        if firstname and lastname:
            amount_paid = formField_list[5].get()
            if int(amount_paid) < total_cost:
                balance = total_cost - int(amount_paid)
                # Todo student id shld be required
                # TODO: Optimize formField DS
                student_id = formField_list[2].get()
                student_address = formField_list[3].get()
                phone_number = formField_list[4].get()
                # amount_paid = formField_list[5].get()
                # Course info
                term_period = formField_list[6].get()
                class_type = formField_list[7].get()
                class_number = formField_list[8].get()
                print("---------------Student Data---------------")
                print("First name:", firstname, "Last name: ", lastname, "\nAmount Paid: ",amount_paid)
                print("ID: ", student_id, "\nAddress: ", student_address, "\nPhone Number: ", phone_number)
                print("Term Period: ", term_period, "\nClass Type: ", class_type, "\nLevel: ", class_number, "\nBalance: ", balance)
                # print("Registration status", registration_status)
                print("------------------------------------------")
                # make sure to update the path syntax for other OS
                filepath = cwd+"\\data.xlsx"
                
                if not os.path.exists(filepath):
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    heading = ["ID", "First Name", "Last Name", "Phone Number", "Address","Term Period", "Class Type", "Level", "Amount Paid", "Balance Due"]
                    sheet.append(heading)
                    workbook.save(filepath)
                workbook = openpyxl.load_workbook(filepath)
                sheet = workbook.active
                sheet.append([student_id, firstname, lastname, phone_number, student_address, term_period, class_type, class_number, amount_paid, balance])
                workbook.save(filepath)

            elif int(amount_paid) > total_cost:
                tkinter.messagebox.showwarning(title="Error", message="Amount cannot be more than 5000.")
                # os._exit(1)
                
        else:
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:
        tkinter.messagebox.showwarning(title= "Error", message="You have not accepted the terms")

def clear_form():
    # TODO: Optimize the logic in this function
    formField_list[0].delete(0, tkinter.END)
    formField_list[1].delete(0, tkinter.END)
    formField_list[2].delete(0, tkinter.END)
    formField_list[3].delete(0, tkinter.END)
    formField_list[4].delete(0, tkinter.END)
    formField_list[5].delete(0, tkinter.END)
    formField_list[6].delete(0, tkinter.END)
    formField_list[7].delete(0, tkinter.END)
    formField_list[8].delete(0, tkinter.END)


def set_form_label_frame(parentFrame, frameTxt=None, r=0, c=0):
    form_label_frame =tkinter.LabelFrame(parentFrame, text=frameTxt)
    form_label_frame.grid(row= r, column=c, padx=20, pady=10)
    return form_label_frame

def set_form_label(labelFrame, labelTxt, r, c):
    form_label = tkinter.Label(labelFrame, text=labelTxt)
    form_label.grid(row=r, column=c)
    return form_label

def set_form_entry(entryFrame, r, c):
    form_entry = tkinter.Entry(entryFrame)
    form_entry.grid(row=r, column=c)
    return form_entry

def set_form_spinbox(spinboxFrame, strt, end, r, c):
    form_spinbox = tkinter.Spinbox(spinboxFrame, from_=strt, to=end)
    form_spinbox.grid(row=r, column=c)
    return form_spinbox

def set_form_combobox(comboboxFrame, val1, val2, val3, r,c):
    form_combobox = ttk.Combobox(comboboxFrame, values=["", val1, val2, val3])
    form_combobox.grid(row=r, column=c)
    return form_combobox

def create_button(frame, btnType, func, r, c):
     # Button
    button = tkinter.Button(frame, text=btnType, command=func, font=("Helvetica", 14))
    button.grid(row=r, column=c, sticky="news", padx=250, pady=5)
    return button

# Creating windows Forms
def setup_form(frame):
        # Saving User Info
    student_info_frame = set_form_label_frame(frame, "Student Information")
    set_form_label(student_info_frame,"First Name", 0, 0)
    first_name_entry = set_form_entry(student_info_frame, 0, 1)
    formField_list.append(first_name_entry)

    set_form_label(student_info_frame,"Last Name", 1,0)
    last_name_entry = set_form_entry(student_info_frame, 1,1)
    formField_list.append(last_name_entry)

    set_form_label(student_info_frame,"Student ID", 2, 0)
    student_id_entry = set_form_entry(student_info_frame, 2,1)
    formField_list.append(student_id_entry)

    set_form_label(student_info_frame,"Address", 3,0)
    student_address_entry = set_form_entry(student_info_frame, 3,1)
    formField_list.append(student_address_entry)

    set_form_label(student_info_frame,"Phone Number", 4, 0) 
    phone_number_entry = set_form_entry(student_info_frame, 4,1)
    formField_list.append(phone_number_entry)

    set_form_label(student_info_frame,"Amount Paid", 5, 0)
    amount_paid_entry = set_form_entry(student_info_frame, 5,1)
    formField_list.append(amount_paid_entry)


    for widget in student_info_frame.winfo_children():
        widget.grid_configure(padx=10, pady=5)

    register_status_frame = set_form_label_frame(frame,"Registration Status",1,0)


    set_form_label(register_status_frame, "Term", 0,0)
    term_spinbox = set_form_spinbox(register_status_frame,1,3, 1, 0)
    formField_list.append(term_spinbox)

    set_form_label(register_status_frame, "Class Type", 0,2)
    class_type_combobox = set_form_combobox(register_status_frame, "class","JSS","SSS",1,2)
    formField_list.append(class_type_combobox)

    set_form_label(register_status_frame, "Level", 0, 3)
    class_number_spinbox = set_form_spinbox(register_status_frame,1,6, 1, 3)
    formField_list.append(class_number_spinbox)

    for widget in register_status_frame.winfo_children():
        widget.grid_configure(padx=10, pady=5)

    terms_frame = set_form_label_frame(frame, "Terms & Conditions", 2,0)
    # terms_frame.place(relx=0.1, rely=0.9)

    accept_var = tkinter.StringVar(value="Not Accepted")
    terms_check = tkinter.Checkbutton(terms_frame, text= "I accept the terms and conditions.",
                                    variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
    terms_check.grid(row=0, column=0)
    '''TODO: Change the formField to an object'''
    formField_list.append(accept_var)

    create_button(frame,btnType="Submit Data", func=submit_data, r=5, c=0)
    clear_btn = create_button(frame,btnType="Clear Entry", func=clear_form, r=5, c=1)
    clear_btn.place(relx=0.2,rely=0.9) # move botton position x-axis=0.2 and y-axis=0.9


def main():
    # code executions
    window = tkinter.Tk()
    window.title("Data Entry Form")
    frame = tkinter.Frame(window)
    frame.pack()
    setup_form(frame)

    window.mainloop()

if __name__ == "__main__":
    main()