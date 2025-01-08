import os
import csv
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Entry, Button, OptionMenu, Text, scrolledtext, Frame, Scrollbar
import tkinter.font as tkfont
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
    # Handle newlines and tabs properly
    lines = []
    for line in text.split('\n'):
        # Skip empty lines but preserve them in output
        if not line.strip():
            lines.append('')
            continue
        # Handle line wrapping
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
    """Process escape sequences in text."""
    try:
        # Handle common escape sequences
        replacements = {
            '\\n': '\n',
            '\\t': '\t',
            '\\r': '\r'
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text
    except:
        return text


def generate_pdfs(template_path, csv_path, output_dir, filename_prefix, custom_text, font_name, font_size, x_percent, y_percent, max_chars):
    try:
        if not template_path or not csv_path or not output_dir or not custom_text:
            raise ValueError("All fields are required!")

        x = float(x_percent) / 100
        y = float(y_percent) / 100
        font_size = int(font_size)
        max_chars = int(max_chars)

        # Read the template PDF
        pdf_reader = PdfReader(template_path)

        with open(csv_path, mode="r", encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            headers = reader.fieldnames  # Dynamically retrieve CSV column names

            for row in reader:
                # Replace tags with values from the current CSV row
                try:
                    output_text = custom_text.format(**row)
                except KeyError as e:
                    raise ValueError(f"Tag '{e.args[0]}' not found in CSV headers: {headers}")

                # Process escape sequences in the text
                output_text = process_escape_sequences(output_text)
                wrapped_lines = wrap_text(output_text, max_chars)

                # Define the output file name
                name_column = headers[0]  # First column as the "name"
                surname_column = headers[1]  # Second column as the "surname"
                output_pdf_path = os.path.join(output_dir, f"{filename_prefix} {row[name_column]} {row[surname_column]}.pdf")

                # Create a new PDF to overlay the text
                packet = BytesIO()
                c = canvas.Canvas(packet, pagesize=letter)
                
                # Set font with Unicode support
                c.setFont(font_name, font_size)

                # Calculate starting Y position from bottom of page
                y_pos = y * letter[1]  # Y position from bottom
                # Adjust for number of lines to center text block
                total_height = len(wrapped_lines) * font_size * 1.2
                y_pos = y_pos + total_height  # Move up by text block height
                
                for line in wrapped_lines:
                    if line.strip():  # Only adjust position for non-empty lines
                        # Handle text encoding for PDF
                        text_to_write = line.strip().encode('utf-8', errors='ignore').decode('utf-8')
                        c.drawString(x * letter[0], y_pos, text_to_write)
                    y_pos -= font_size * 1.2  # Move down for next line

                c.save()

                # Merge the overlay with the template
                packet.seek(0)
                overlay_pdf = PdfReader(packet)
                output_pdf = PdfWriter()

                for page in pdf_reader.pages:
                    page.merge_page(overlay_pdf.pages[0])
                    output_pdf.add_page(page)

                with open(output_pdf_path, "wb") as out_file:
                    output_pdf.write(out_file)

        messagebox.showinfo("Success", "PDFs generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def update_csv_headers(csv_path_var, headers_var):
    try:
        csv_path = csv_path_var.get()
        if not csv_path:
            raise ValueError("Please select a CSV file first.")
        with open(csv_path, mode="r", encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            headers = reader.fieldnames
            headers_var.set(f"Available tags: {', '.join(headers)}")
    except Exception as e:
        headers_var.set(f"Error: {str(e)}")


def main():
    # Initialize root with proper input method support
    import sys
    if sys.platform.startswith('win'):
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    
    root = Tk()
    root.title("PDF Generator")
    
    # Configure Tkinter for proper text input
    import locale
    try:
        # Set UTF-8 encoding for Tcl/Tk
        root.tk.call('encoding', 'system', 'utf-8')
        
        # Configure locale and encoding
        if sys.platform.startswith('win'):
            # Windows-specific configuration
            import ctypes
            
            # Set console code page to UTF-8
            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleCP(65001)
            kernel32.SetConsoleOutputCP(65001)
            
            # Try multiple locales
            for loc in ['en_US.UTF-8', 'English_United States.1252', '']:
                try:
                    locale.setlocale(locale.LC_ALL, loc)
                    break
                except locale.Error:
                    continue
        else:
            # Unix-like systems
            locale.setlocale(locale.LC_ALL, '')
            os.environ['LANG'] = 'en_US.UTF-8'
            os.environ['LC_ALL'] = 'en_US.UTF-8'
            
        # Configure Tcl/Tk for proper text handling
        root.tk.eval('''
            encoding system utf-8
            fconfigure stdin -encoding utf-8
            fconfigure stdout -encoding utf-8
            fconfigure stderr -encoding utf-8
        ''')
        
    except Exception as e:
        print(f"Warning: Could not fully configure text input: {e}")
    
    # Register Unicode-compatible font
    try:
        pdfmetrics.registerFont(TTFont('DejaVuSans', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
        default_font = 'DejaVuSans'
    except:
        default_font = 'Helvetica'  # Fallback to built-in font

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
        text = text_widget.get("1.0", "end-1c")
        return process_escape_sequences(text)

    Label(root, text="Template PDF:").grid(row=0, column=0)
    Entry(root, textvariable=template_var, width=50).grid(row=0, column=1)
    Button(root, text="Browse", command=lambda: select_file(template_var, "PDF")).grid(row=0, column=2)

    Label(root, text="CSV File:").grid(row=1, column=0)
    Entry(root, textvariable=csv_var, width=50).grid(row=1, column=1)
    Button(root, text="Browse", command=lambda: [select_file(csv_var, "CSV"), update_csv_headers(csv_var, headers_var)]).grid(row=1, column=2)

    Label(root, text="Output Directory:").grid(row=2, column=0)
    Entry(root, textvariable=output_dir_var, width=50).grid(row=2, column=1)
    Button(root, text="Browse", command=lambda: select_directory(output_dir_var)).grid(row=2, column=2)

    Label(root, text="Custom Text (use CSV column names as tags, type \\n for newline, \\t for tab):").grid(row=3, column=0)
    # Create text widget with proper Unicode configuration
    text_frame = Frame(root)
    text_frame.grid(row=3, column=1, sticky='nsew')
    
    # Create text widget with comprehensive Unicode support
    text_widget = Text(text_frame, width=38, height=4,
                      wrap='word', undo=True,
                      font='TkDefaultFont',
                      insertwidth=2,  # Make cursor more visible
                      maxundo=0,  # Unlimited undo
                      exportselection=1,  # Allow copy/paste
                      insertofftime=0,  # Always show cursor
                      blockcursor=True if sys.platform.startswith('win') else False,
                      spacing1=2, spacing2=2,  # Add some padding
                      insertborderwidth=3,
                      tabs=('1c',))  # Set proper tab stops
    
    # Configure additional text widget properties
    text_widget.configure(background='white')  # Ensure consistent background
    text_widget.configure(selectbackground='#0078d7')  # Better selection visibility
    text_widget.configure(relief='sunken')  # Better visual feedback
    
    # Configure text widget for better input handling
    text_widget.configure(inactiveselectbackground=text_widget.cget('selectbackground'))
    
    # Configure text widget for Unicode input
    if sys.platform.startswith('win'):
        text_widget.configure(imemode='active')
    text_widget.pack(side='left', fill='both', expand=True)
    
    # Add scrollbar
    scrollbar = Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
    scrollbar.pack(side='right', fill='y')
    text_widget.configure(yscrollcommand=scrollbar.set)
    
    # Add text change handler for escape sequences
    def on_text_change(event=None):
        if not text_widget.edit_modified():  # Skip if this callback caused the change
            return
        try:
            content = text_widget.get("1.0", "end-1c")
            if '\\n' in content or '\\t' in content:
                # Get cursor position and selection
                insert_pos = text_widget.index("insert")
                try:
                    sel_start = text_widget.index("sel.first")
                    sel_end = text_widget.index("sel.last")
                    has_selection = True
                except:
                    has_selection = False
                
                # Process escape sequences
                new_content = process_escape_sequences(content)
                
                # Update content if changed
                if new_content != content:
                    text_widget.delete("1.0", "end")
                    text_widget.insert("1.0", new_content)
                    
                    # Restore cursor and selection
                    text_widget.mark_set("insert", insert_pos)
                    if has_selection:
                        text_widget.tag_add("sel", sel_start, sel_end)
        except:
            pass  # Keep original text if processing fails
        finally:
            text_widget.edit_modified(False)
            
    text_widget.bind('<<Modified>>', on_text_change)
    Label(root, textvariable=headers_var, wraplength=400, justify="left").grid(row=4, column=1)

    Label(root, text="Font:").grid(row=5, column=0)
    font_options = [default_font, "Helvetica", "Times-Roman", "Courier"]
    font_var.set(default_font)
    OptionMenu(root, font_var, *font_options).grid(row=5, column=1)

    Label(root, text="Font Size:").grid(row=6, column=0)
    Entry(root, textvariable=font_size_var).grid(row=6, column=1)

    Label(root, text="X Position (%):").grid(row=7, column=0)
    Entry(root, textvariable=x_percent_var).grid(row=7, column=1)

    Label(root, text="Y Position (%):").grid(row=8, column=0)
    Entry(root, textvariable=y_percent_var).grid(row=8, column=1)

    Label(root, text="Max Characters per Line:").grid(row=9, column=0)
    Entry(root, textvariable=max_chars_var).grid(row=9, column=1)

    Label(root, text="Filename Prefix:").grid(row=10, column=0)
    Entry(root, textvariable=filename_prefix_var).grid(row=10, column=1)

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
    )).grid(row=11, column=1)

    root.mainloop()


if __name__ == "__main__":
    main()
