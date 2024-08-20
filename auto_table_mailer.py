import tkinter as tk
from tkinter import ttk
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import win32com.client as win32
import tkinter.messagebox as messagebox

def calculate():
    # This function will update the table view based on the input values
    values = [
        line1_entry.get(), line2_entry.get(), line3_entry.get(), line4_entry.get(),
        line5_combobox.get(), line6_entry.get(), line7_entry.get(), line8_entry.get(),
        line9_entry.get()  # Get the value from the line9 entry
    ]
    # Get the current number of items in the table
    num_items = len(table.get_children())
    # Determine the tag for the new row
    row_tags = ("oddrow",) if num_items % 2 == 0 else ("evenrow",)
    table.insert("", "end", values=values, tags=row_tags)

def focus_next_widget(event):
    event.widget.tk_focusNext().focus()
    return "break"

def handle_return(event):
    if event.widget == calculate_button:
        calculate()
        root.focus_set()  # Move focus back to root to enable global binding
        return "break"  # Return to prevent further processing
    else:
        focus_next_widget(event)
        return "break"

def send_email():
    # Connect to Outlook
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)

    # Construct email content in HTML format
    html_content = "<html><body>"
    html_content += "<p>Hi,</p>"
    html_content += "<br>"  # Add a blank line    
    html_content += "<p>First line of message.</p>"
    html_content += "<br>"  # Add a blank line
    html_content += "<table border='1' style='text-align:center; width:60%; padding: 5px;'><thead style='background-color: lightblue;'><tr>"
    for field in fields:
        html_content += f"<th>{field}</th>"
    html_content += "</tr></thead><tbody>"

    # Add rows with alternating background colors
    for index, item in enumerate(table.get_children()):
        if index % 2 == 0:
            html_content += "<tr>"
        else:
            html_content += "<tr style='background-color: lightgrey;'>"
        for value in table.item(item)['values']:
            html_content += f"<td>{value}</td>"
        html_content += "</tr>"
    
    html_content += "</tbody></table>"
    html_content += "<br>"  # Add a blank line
    html_content += "<p>Regards,</p>"
    #html_content += "<br>"  # Add a blank line
    html_content += "<p style='font-family: Georgia, serif; font-size: 11pt; color: #022a4d; margin-bottom: 0;'><strong>First Name</strong><br><span style='font-size: 8pt;'>Job Title</span></p>"
    html_content += "<p style='font-family: Georgia, serif; font-size: 11pt; color: #022a4d; margin-bottom: 5px;'><strong>Company Name</strong><br>"
    html_content += "Company Name, Street Name, City, Post Code<br>"
    html_content += "Tel Number<br>"
    html_content += "Fax Number<br>"
    html_content += "Website Address<br>"
    html_content += "Email Address</p>"
    html_content += "</body></html>"

    # Set email properties
    mail.Subject = "Subject Here"
    mail.HTMLBody = html_content

    # Add recipients - Using selected email addresses from check buttons
    recipients = [email_options[i] for i, value in enumerate(selected_emails) if value.get() == 1]
    mail.To = ";".join(recipients)

    # Send email
    try:
        mail.Send()
        messagebox.showinfo("Success", "Email sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email: {str(e)}")

root = tk.Tk()
root.title("Info to Email")

# Frame for input fields
input_frame = tk.Frame(root)
input_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

# Labels and Entry Widgets
padx_labels = (20, 10)
padx_entries = (0, 5)

line1_label = tk.Label(input_frame, text="line1")
line1_label.grid(row=0, column=0, sticky="w", padx=padx_labels, pady=5)
line1_entry = tk.Entry(input_frame)
line1_entry.grid(row=0, column=1, padx=padx_entries, pady=5)
line1_entry.bind("<Return>", focus_next_widget)

line2_label = tk.Label(input_frame, text="line2")
line2_label.grid(row=1, column=0, sticky="w", padx=padx_labels, pady=5)
line2_entry = tk.Entry(input_frame)
line2_entry.grid(row=1, column=1, padx=padx_entries, pady=5)
line2_entry.bind("<Return>", focus_next_widget)

line3_label = tk.Label(input_frame, text="line3")
line3_label.grid(row=2, column=0, sticky="w", padx=padx_labels, pady=5)
line3_entry = tk.Entry(input_frame)
line3_entry.grid(row=2, column=1, padx=padx_entries, pady=5)
line3_entry.bind("<Return>", focus_next_widget)

line4_label = tk.Label(input_frame, text="line4")
line4_label.grid(row=3, column=0, sticky="w", padx=padx_labels, pady=5)
line4_entry = tk.Entry(input_frame)
line4_entry.grid(row=3, column=1, padx=padx_entries, pady=5)
line4_entry.bind("<Return>", focus_next_widget)

line5_label = tk.Label(input_frame, text="line5")
line5_label.grid(row=4, column=0, sticky="w", padx=padx_labels, pady=5)
line5_combobox = ttk.Combobox(input_frame, values=['example1', 'example2', 'example3'], width=17)  # Adjust width here
line5_combobox.grid(row=4, column=1, padx=padx_entries, pady=5)
line5_combobox.bind("<Return>", focus_next_widget)

line6_label = tk.Label(input_frame, text="line6")
line6_label.grid(row=5, column=0, sticky="w", padx=padx_labels, pady=5)
line6_entry = tk.Entry(input_frame)
line6_entry.grid(row=5, column=1, padx=padx_entries, pady=5)
line6_entry.bind("<Return>", focus_next_widget)

line7_label = tk.Label(input_frame, text="line7")
line7_label.grid(row=6, column=0, sticky="w", padx=padx_labels, pady=5)
line7_entry = tk.Entry(input_frame)
line7_entry.grid(row=6, column=1, padx=padx_entries, pady=5)
line7_entry.bind("<Return>", focus_next_widget)

line8_label = tk.Label(input_frame, text="line8")
line8_label.grid(row=7, column=0, sticky="w", padx=padx_labels, pady=5)
line8_entry = tk.Entry(input_frame)
line8_entry.grid(row=7, column=1, padx=padx_entries, pady=5)
line8_entry.bind("<Return>", focus_next_widget)

# Add line9 Entry
line9_label = tk.Label(input_frame, text="line9")
line9_label.grid(row=8, column=0, sticky="w", padx=padx_labels, pady=5)
line9_entry = tk.Entry(input_frame)
line9_entry.grid(row=8, column=1, padx=padx_entries, pady=5)
line9_entry.bind("<Return>", focus_next_widget)

# Check buttons for selecting email addresses
email_label = tk.Label(input_frame, text="Select Email Addresses:")
email_label.grid(row=9, column=0, sticky="w", padx=padx_labels, pady=5)

email_options = ['name1@example.com', 'name2@example.com', 'name3@example.com']
selected_emails = [tk.IntVar() for _ in email_options]
for i, email in enumerate(email_options):
    email_check = tk.Checkbutton(input_frame, text=email, variable=selected_emails[i], onvalue=1, offvalue=0)
    email_check.grid(row=9, column=i+1, sticky="w", padx=5)

# Calculate Button
calculate_button = tk.Button(input_frame, text="Update", command=calculate)
calculate_button.grid(row=15, columnspan=2, pady=10)
calculate_button.bind("<Return>", handle_return)

# Send Email Button
send_email_button = tk.Button(input_frame, text="Send Email", command=send_email)
send_email_button.grid(row=16, columnspan=2, pady=10)

# Frame for table view
table_frame = tk.Frame(root)
table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Define fields list for table view
fields = ["line1", "line2", "line3", "line4", "line5", "line6", "line7", "line8", "line9"]

# Treeview style
style = ttk.Style()
style.theme_use("default")
style.configure("Treeview", background="#D3D3D3", foreground="black", rowheight=25, fieldbackground="white")
style.map("Treeview", background=[("selected", "#347083")])

table = ttk.Treeview(table_frame, columns=fields, show="headings", style="Treeview")
for field in fields:
    table.heading(field, text=field, anchor="center")
    table.column(field, anchor="center", width=100)  # Set custom width for each column

# Alternate row colors
table.tag_configure('oddrow', background='white')
table.tag_configure('evenrow', background='#D3D3D3')

# Bind the Return key to the calculate function globally
root.bind("<Return>", handle_return)

table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

root.mainloop()
