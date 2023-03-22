import customtkinter as tk
from tkinter import ttk
import openai
from apikey import API_KEY
from tkinter import filedialog
from tkinter import messagebox
import os
import docx
from fpdf import FPDF

pdf = FPDF(
    "P",
    "mm",
    "A4",
)
tk.set_appearance_mode("dark")
tk.set_default_color_theme("dark-blue")


class App:
    def __init__(self, master):
        self.master = master
        self.file_path = ""
        self.summary = ""
        self.file_path_list = []
        self.heading_options = []

        def heading_dropdown_callback(choice):
            self.subheading_dropdown.configure(state="normal")
            self.subheading_dropdown.set("Select a subheading")
            print("combobox dropdown clicked:", choice)

        # create select file button
        self.select_file_button = tk.CTkButton(
            self.master, text="Select File", command=self.select_file
        )
        self.select_file_button.grid(row=0, column=3)

        # Configure 6 rows and 4 columns
        for i in range(16):
            root.rowconfigure(i, weight=1)
        for i in range(7):
            root.columnconfigure(i, weight=1)

        # create label to show selected file path
        self.file_path_label = tk.CTkLabel(self.master, text="No file selected.")
        self.file_path_label.grid(row=1, column=1, columnspan=5, sticky="ew")

        # Create the first Combobox and place it on the left side
        self.heading_dropdown = tk.CTkComboBox(
            self.master,
            values=self.heading_options,
            command=heading_dropdown_callback,
            state="disabled",
        )
        self.heading_dropdown.grid(row=2, column=1, columnspan=2, sticky="ew")

        # Create the second Combobox and place it on the right side
        self.subheading_dropdown = self.subheading_dropdown = tk.CTkComboBox(
            self.master,
            values=self.heading_options,
            state="disabled",
        )
        self.heading_dropdown.set("Select a Subheading")
        self.subheading_dropdown.grid(row=2, column=4, columnspan=2, sticky="ew")

        # create button to show file content
        self.show_content_button = tk.CTkButton(
            self.master, text="Show Summary", command=self.show_summary
        )
        self.show_content_button.grid(row=3, column=3)

        # create text area to show file content
        self.content_text = tk.CTkTextbox(self.master)
        self.content_text.grid(row=4, column=1, columnspan=5, rowspan=8, sticky="nsew")

        master.title("File Save")

        # Create a button to save output as PDF
        self.pdf_button = tk.CTkButton(
            self.master, text="Save as PDF", command=self.save_as_pdf
        )
        self.pdf_button.grid(row=13, column=1, columnspan=2, sticky="ew")

        # Create a button to save output as TXT
        self.txt_button = tk.CTkButton(
            self.master, text="Save as TXT", command=self.save_as_txt
        )
        self.txt_button.grid(row=13, column=4, columnspan=2, sticky="ew")

    def save_as_pdf(self):
        # Add a page to the PDF object
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        text = self.multiline_it(self.summary)
        print(text)

        # Set the cell width and write the text to the cell
        for i in text:
            pdf.cell(0, 10, i, ln=1)

        # Save the PDF file to the downloads folder
        pdf_file = os.path.join(os.path.expanduser("~"), "Downloads", "output.pdf")
        pdf.output(pdf_file)

    # helper function for pdf creator
    def multiline_it(self, stri):
        stri = stri.split()
        result = [""]
        chars = 80
        c = 0
        for i in stri:
            c += len(i)
            if c > chars:
                result[-1] += i
                result.append("")
                c = 0
            else:
                result[-1] += i + " "
        return result

    # Save the text file to the downloads folder
    def save_as_txt(self):
        txt_file = os.path.join(os.path.expanduser("~"), "Downloads", "output.txt")
        with open(txt_file, "w") as f:
            f.write(self.summary)

    # File selection dialog
    def select_file(self):
        self.file_path = filedialog.askopenfilename()
        self.file_path_label.configure(text=self.file_path)
        self.file_path_list = self.file_path.split(".")
        ext = self.file_path_list[-1]
        if ext == "docx":
            self.heading_dropdown.configure(state="normal")
            self.heading_dropdown.set("Select a Heading")
            doc = docx.Document(self.file_path)
            # print(doc)
            for para in doc.paragraphs:
                if para.style.name == "Heading 1":
                    self.heading_options.append(para.text)
            self.heading_dropdown.configure(values=self.heading_options)
        elif ext == "pdf":
            pass
        else:
            pass

    # API Call and summary texbox
    def show_summary(self):
        if self.file_path:
            contents = ""
            with open(self.file_path, "r") as file:
                contents = file.read()

            contents += "\n Summarize it."

            try:
                openai.api_key = API_KEY
                completion = openai.Completion.create(
                    engine="text-davinci-003", prompt=contents, max_tokens=200
                )
                # print(completion)
                self.summary = completion.choices[0]["text"].strip()

                print("\nSummary of the file content :\n" + self.summary + "\n")
                self.content_text.delete("1.0", tk.END)
                self.content_text.insert(tk.END, self.summary)

            except:
                messagebox.showerror(
                    "Error", "Error: Your file content is exceeding the input limit!!!"
                )
        else:
            self.content_text.delete("1.0", tk.END)
            messagebox.showerror("Error", "Please select a file first.")


root = tk.CTk()
root.title("File Summarizer")
root.geometry("800x500")
app = App(root)
root.mainloop()
