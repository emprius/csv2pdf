import os
import csv
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Entry, Button, OptionMenu, Frame, Scrollbar, Text, font
from tkinter.ttk import Button as TtkButton, Style
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO


def select_file(file_var, file_type):
    filetypes = [("PDF files", "*.pdf")] if file_type == "PDF" else [("CSV files", "*.csv")]
    filepath = filedialog.askopenfilename(filetypes=filetypes)
    file_var.set(filepath)


def select_directory(directory_var):
    directory = filedialog.askdirectory()
    directory_var.set(directory)


def wrap_text(text, max_chars):
    lines = []
    for line in text.split('\n'):
        if not line.strip():
            lines.append('')
            continue
        current_line = ''
        words = line.split()
        for word in words:
            if not current_line:
                current_line = word
            elif len(current_line) + 1 + len(word) <= max_chars:
                current_line += ' ' + word
            else:
                lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
    return lines


def process_escape_sequences(text):
    replacements = {
        '\\n': '\n',
        '\\t': '\t',
        '\\r': '\r'
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def generate_pdfs(template_path, csv_path, output_dir, filename_prefix, custom_text, font_name, font_size, x_percent, y_percent, max_chars):
    try:
        if not template_path or not csv_path or not output_dir or not custom_text:
            raise ValueError("All fields are required!")

        x = float(x_percent) / 100
        y = float(y_percent) / 100
        font_size = int(font_size)
        max_chars = int(max_chars)

        pdf_reader = PdfReader(template_path)

        with open(csv_path, mode="r", encoding='utf-8-sig') as csvfile:  # utf-8-sig handles BOM character
            reader = csv.DictReader(csvfile)
            headers = reader.fieldnames

            for row in reader:
                try:
                    # Format text with row data
                    formatted_text = custom_text
                    for key, value in row.items():
                        placeholder = "{" + key + "}"
                        formatted_text = formatted_text.replace(placeholder, value)
                    
                    output_text = process_escape_sequences(formatted_text)
                    wrapped_lines = wrap_text(output_text, max_chars)

                    name_column = headers[0]
                    surname_column = headers[1]
                    output_pdf_path = os.path.join(output_dir, f"{filename_prefix} {row[name_column]} {row[surname_column]}.pdf")

                    # Create a new PDF with the text
                    packet = BytesIO()
                    c = canvas.Canvas(packet, pagesize=letter)
                    c.setFont(font_name, font_size)

                    # Calculate starting position from bottom
                    y_pos = y * letter[1]
                    line_height = font_size * 1.2  # Line spacing
                    total_height = len(wrapped_lines) * line_height
                    y_pos = y_pos + total_height  # Start from top of text block

                    # Draw text
                    for line in wrapped_lines:
                        if line.strip():
                            text = line.strip()
                            c.drawString(x * letter[0], y_pos, text)
                        y_pos -= line_height

                    c.save()
                    packet.seek(0)

                    # Create output PDF
                    output_pdf = PdfWriter()
                    # Create a fresh copy of the template page
                    template_page = PdfReader(template_path).pages[0]
                    
                    # Create overlay
                    overlay = PdfReader(packet)
                    template_page.merge_page(overlay.pages[0])
                    
                    # Add merged page to output
                    output_pdf.add_page(template_page)

                    # Write the output file
                    with open(output_pdf_path, "wb") as out_file:
                        output_pdf.write(out_file)
                except KeyError as e:
                    raise ValueError(f"Tag '{e.args[0]}' not found in CSV headers: {headers}")

        messagebox.showinfo("Success", "PDFs generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def update_csv_headers(csv_path_var, headers_var):
    try:
        csv_path = csv_path_var.get()
        if not csv_path:
            raise ValueError("Please select a CSV file first.")
        with open(csv_path, mode="r", encoding='utf-8-sig') as csvfile:  # utf-8-sig handles BOM character
            reader = csv.DictReader(csvfile)
            headers = reader.fieldnames
            headers_var.set(f"Available tags: {', '.join(headers)}")
    except Exception as e:
        headers_var.set(f"Error: {str(e)}")


def main():
    root = Tk()
    root.title("PDF Generator")
    
    # Configure ttk style for toolbar buttons
    style = Style()
    style.configure('Toolbar.TButton', padding=4)

    template_var = StringVar()
    csv_var = StringVar()
    output_dir_var = StringVar()
    filename_prefix_var = StringVar()
    font_var = StringVar()
    font_size_var = StringVar(value="12")
    x_percent_var = StringVar(value="10")
    y_percent_var = StringVar(value="10")
    max_chars_var = StringVar(value="50")
    headers_var = StringVar(value="Available tags: None (select a CSV file)")

    def get_custom_text():
        return text_widget.get("1.0", "end-1c")

    Label(root, text="Template PDF:").grid(row=0, column=0)
    Entry(root, textvariable=template_var, width=50).grid(row=0, column=1)
    Button(root, text="Browse", command=lambda: select_file(template_var, "PDF")).grid(row=0, column=2)

    Label(root, text="CSV File:").grid(row=1, column=0)
    Entry(root, textvariable=csv_var, width=50).grid(row=1, column=1)
    Button(root, text="Browse", command=lambda: [select_file(csv_var, "CSV"), update_csv_headers(csv_var, headers_var)]).grid(row=1, column=2)

    Label(root, text="Output Directory:").grid(row=2, column=0)
    Entry(root, textvariable=output_dir_var, width=50).grid(row=2, column=1)
    Button(root, text="Browse", command=lambda: select_directory(output_dir_var)).grid(row=2, column=2)

    Label(root, text="Custom Text (use CSV column names as tags):").grid(row=3, column=0)
    
    # Create a frame for the toolbar
    toolbar_frame = Frame(root)
    toolbar_frame.grid(row=3, column=1, sticky='ew')
    
    # Create text frame with scrollbar
    text_frame = Frame(root)
    text_frame.grid(row=4, column=1, sticky='nsew')
    
    # Configure text widget with UTF-8 and IME support
    text_widget = Text(text_frame, width=50, height=10, wrap='word', undo=True)
    text_widget.configure(font=('TkDefaultFont', 10))  # Use system default font which supports Unicode
    scrollbar = Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)
    
    # Enable input method support
    def handle_keypress(event):
        # Allow all key events to be processed normally
        return None
    
    text_widget.bind('<Key>', handle_keypress)
    text_widget.bind('<KeyPress>', handle_keypress)
    text_widget.bind('<KeyRelease>', handle_keypress)
    
    # Pack the text widget and scrollbar
    text_widget.pack(side='left', fill='both', expand=True)
    scrollbar.pack(side='right', fill='y')
    
    # Define text formatting functions
    def apply_bold():
        try:
            current_tags = text_widget.tag_names("sel.first")
            if "bold" in current_tags:
                text_widget.tag_remove("bold", "sel.first", "sel.last")
            else:
                text_widget.tag_add("bold", "sel.first", "sel.last")
                text_widget.tag_configure("bold", font=font.Font(weight="bold"))
        except:
            messagebox.showinfo("Info", "Please select text to format")
    
    def apply_italic():
        try:
            current_tags = text_widget.tag_names("sel.first")
            if "italic" in current_tags:
                text_widget.tag_remove("italic", "sel.first", "sel.last")
            else:
                text_widget.tag_add("italic", "sel.first", "sel.last")
                text_widget.tag_configure("italic", font=font.Font(slant="italic"))
        except:
            messagebox.showinfo("Info", "Please select text to format")
    
    def apply_underline():
        try:
            current_tags = text_widget.tag_names("sel.first")
            if "underline" in current_tags:
                text_widget.tag_remove("underline", "sel.first", "sel.last")
            else:
                text_widget.tag_add("underline", "sel.first", "sel.last")
                text_widget.tag_configure("underline", underline=True)
        except:
            messagebox.showinfo("Info", "Please select text to format")
    
    # Create formatting buttons with custom fonts and tooltips
    bold_font = font.Font(weight='bold')
    italic_font = font.Font(slant='italic')
    underline_font = font.Font(underline=True)
    
    bold_button = TtkButton(toolbar_frame, text="B", width=3, command=apply_bold, style='Toolbar.TButton')
    italic_button = TtkButton(toolbar_frame, text="I", width=3, command=apply_italic, style='Toolbar.TButton')
    underline_button = TtkButton(toolbar_frame, text="U", width=3, command=apply_underline, style='Toolbar.TButton')
    
    # Configure button fonts
    bold_button.configure(text="B")
    italic_button.configure(text="I")
    underline_button.configure(text="U")
    
    # Add tooltips
    def create_tooltip(widget, text):
        def show_tooltip(event):
            tooltip = Label(root, text=text, relief="solid", borderwidth=1)
            tooltip.place_forget()
            
            def position_tooltip():
                widget_x = widget.winfo_rootx()
                widget_y = widget.winfo_rooty()
                tooltip.place(x=root.winfo_x() + widget_x, 
                            y=root.winfo_y() + widget_y + widget.winfo_height())
            
            position_tooltip()
            widget.tooltip = tooltip
            
        def hide_tooltip(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind('<Enter>', show_tooltip)
        widget.bind('<Leave>', hide_tooltip)
    
    create_tooltip(bold_button, "Bold (Ctrl+B)")
    create_tooltip(italic_button, "Italic (Ctrl+I)")
    create_tooltip(underline_button, "Underline (Ctrl+U)")
    
    # Pack formatting buttons with improved spacing
    bold_button.pack(side='left', padx=4)
    italic_button.pack(side='left', padx=4)
    underline_button.pack(side='left', padx=4)
    
    # Add keyboard shortcuts without interfering with IME
    def handle_shortcuts(event):
        if event.state & 4 and event.keysym in ['b', 'i', 'u']:  # Check for Ctrl key and specific shortcuts
            if event.keysym == 'b':
                apply_bold()
            elif event.keysym == 'i':
                apply_italic()
            elif event.keysym == 'u':
                apply_underline()
            return 'break'
        return None  # Allow all other keys to be processed normally
    
    text_widget.bind('<Key>', handle_shortcuts)
    
    # Clear default text when clicked
    def clear_default_text(event):
        if text_widget.get("1.0", "end-1c") == "Type your text here...":
            text_widget.delete("1.0", "end")
            text_widget.unbind('<Button-1>')
    
    text_widget.insert("1.0", "Type your text here...")
    text_widget.bind('<Button-1>', clear_default_text)

    Label(root, textvariable=headers_var, wraplength=400, justify="left").grid(row=5, column=1)

    Label(root, text="Font:").grid(row=6, column=0)
    font_options = [
        "Helvetica", "Times-Roman", "Courier",
        "Liberation Sans", "DejaVu Sans", "Arial",
        "Ubuntu", "Noto Sans", "FreeSans"
    ]
    font_var.set("Helvetica")  # Default to Helvetica as it's always available
    OptionMenu(root, font_var, *font_options).grid(row=6, column=1)

    Label(root, text="Font Size:").grid(row=7, column=0)
    Entry(root, textvariable=font_size_var).grid(row=7, column=1)

    Label(root, text="X Position (%):").grid(row=8, column=0)
    Entry(root, textvariable=x_percent_var).grid(row=8, column=1)

    Label(root, text="Y Position (%):").grid(row=9, column=0)
    Entry(root, textvariable=y_percent_var).grid(row=9, column=1)

    Label(root, text="Max Characters per Line:").grid(row=10, column=0)
    Entry(root, textvariable=max_chars_var).grid(row=10, column=1)

    Label(root, text="Filename Prefix:").grid(row=11, column=0)
    Entry(root, textvariable=filename_prefix_var).grid(row=11, column=1)

    Button(root, text="Generate PDFs", command=lambda: generate_pdfs(
        template_var.get(),
        csv_var.get(),
        output_dir_var.get(),
        filename_prefix_var.get(),
        get_custom_text(),
        font_var.get(),
        font_size_var.get(),
        x_percent_var.get(),
        y_percent_var.get(),
        max_chars_var.get()
    )).grid(row=12, column=1)

    root.mainloop()


if __name__ == "__main__":
    main()
