import os
import csv
import json
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
    # Split text into words while preserving spaces
    words = []
    current_word = ''
    for char in text:
        if char.isspace():
            if current_word:
                words.append(current_word)
                current_word = ''
            words.append(char)
        else:
            current_word += char
    if current_word:
        words.append(current_word)

    lines = []
    current_line = ''
    
    for word in words:
        # Handle newlines explicitly
        if word == '\n':
            if current_line:
                lines.append(current_line)
                current_line = ''
            lines.append('')
            continue
            
        # Preserve all whitespace at start of line
        if not current_line and word.isspace():
            current_line = word
            continue
            
        # Check if adding word would exceed max_chars
        test_line = current_line + word if current_line else word
        if len(test_line) <= max_chars:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word if not word.isspace() else ''

    if current_line:
        lines.append(current_line)
        
    return lines


def process_escape_sequences(text):
    # Process \n first to handle line breaks
    text = text.replace('\\n', '\n')
    
    # Process \t by adding 4 spaces for each tab
    lines = text.split('\n')
    processed_lines = []
    for line in lines:
        # Process tabs at the beginning of the line first
        while line.startswith('\\t'):
            line = '    ' + line[2:]  # Replace \t with 4 spaces
            
        # Then process any remaining tabs in the line
        while '\\t' in line:
            tab_pos = line.find('\\t')
            line = line[:tab_pos] + '    ' + line[tab_pos+2:]
        processed_lines.append(line)
    
    # Join lines back together
    text = '\n'.join(processed_lines)
    
    # Process other escape sequences
    text = text.replace('\\r', '\r')
    return text


def generate_pdfs(root_window, template_path, csv_path, output_dir, filename_prefix, custom_text, font_name, font_size, x_percent, y_percent, max_chars, text_widget):
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
                    # Get original text
                    original_text = text_widget.get("1.0", "end-1c")
                    
                    # Replace tags first
                    text_with_tags = original_text
                    for key, value in row.items():
                        placeholder = "{" + key + "}"
                        text_with_tags = text_with_tags.replace(placeholder, str(value))
                    
                    # Then process escape sequences
                    processed_text = process_escape_sequences(text_with_tags)
                    
                    # Create a temporary text widget to handle formatting
                    temp_widget = Text(root_window)
                    temp_widget.insert("1.0", processed_text)
                    
                    # Copy formatting from original widget
                    for tag in ["bold", "italic", "underline"]:
                        ranges = text_widget.tag_ranges(tag)
                        for i in range(0, len(ranges), 2):
                            start = ranges[i]
                            end = ranges[i+1]
                            temp_widget.tag_add(tag, start, end)
                    
                    # Get text with formatting by segments
                    formatted_text = []
                    start = "1.0"
                    line_start = True  # Track if we're at the start of a line
                    
                    while True:
                        if temp_widget.compare(start, ">=", "end-1c"):
                            break
                            
                        # Get current formatting
                        formats = []
                        for tag in ["bold", "italic", "underline"]:
                            if tag in temp_widget.tag_names(start):
                                formats.append(tag)
                        
                        # Get current character
                        char = temp_widget.get(start)
                        
                        # Handle line start indentation
                        if line_start and char.isspace():
                            # Find all spaces at start of line
                            next_pos = start
                            spaces = ""
                            while True:
                                if temp_widget.compare(next_pos, ">=", "end-1c"):
                                    break
                                char = temp_widget.get(next_pos)
                                if not char.isspace():
                                    break
                                spaces += char
                                next_pos = temp_widget.index(f"{next_pos}+1c")
                            
                            if spaces:
                                formatted_text.append((spaces, formats))
                                start = next_pos
                                line_start = False
                                continue
                        
                        # Find next format change or space
                        next_pos = start
                        while True:
                            if temp_widget.compare(next_pos, ">=", "end-1c"):
                                break
                            
                            next_char = temp_widget.get(next_pos)
                            next_formats = []
                            for tag in ["bold", "italic", "underline"]:
                                if tag in temp_widget.tag_names(next_pos):
                                    next_formats.append(tag)
                                    
                            # Break if formatting changes or we hit a space
                            if next_formats != formats or next_char.isspace():
                                break
                                
                            next_pos = temp_widget.index(f"{next_pos}+1c")
                        
                        # Get text segment
                        text = temp_widget.get(start, next_pos)
                        if text:
                            # Add segment with its formatting
                            formatted_text.append((text, formats))
                            
                        # If we stopped at a space, add it as a separate segment
                        if temp_widget.compare(next_pos, "<", "end-1c"):
                            space_char = temp_widget.get(next_pos)
                            if space_char.isspace():
                                # Add space as a separate segment without formatting
                                formatted_text.append((space_char, []))
                                next_pos = temp_widget.index(f"{next_pos}+1c")
                                # Check if this is a newline
                                if space_char == '\n':
                                    line_start = True
                            
                        start = next_pos
                        
                    # Clean up temporary widget
                    temp_widget.destroy()
                    
                    # Process text segments
                    processed_text = []
                    for text, formats in formatted_text:
                        processed_text.append((text, formats))
                    
                    # Setup PDF
                    # Format filename with tags
                    filename = filename_prefix.strip()
                    if not filename:  # Handle empty filename
                        filename = "document"
                    
                    # Replace tags in filename
                    try:
                        # First check if all tags exist in headers
                        import re
                        tags = re.findall(r'\{([^}]+)\}', filename)
                        for tag in tags:
                            if tag not in row:
                                raise KeyError(tag)
                        # Then do the replacement
                        filename = filename.format(**row)
                    except KeyError as e:
                        raise ValueError(f"Tag '{e.args[0]}' not found in CSV headers: {headers}")
                    
                    # Handle .pdf extension
                    if not filename.lower().endswith('.pdf'):
                        filename += '.pdf'
                    
                    output_pdf_path = os.path.join(output_dir, filename)
                    
                    packet = BytesIO()
                    c = canvas.Canvas(packet, pagesize=letter)
                    
                    # Set up fonts - use reportlab's built-in fonts
                    if font_name == "Helvetica":
                        base_font = "Helvetica"
                        bold_font = "Helvetica-Bold"
                        italic_font = "Helvetica-Oblique"
                        bold_italic_font = "Helvetica-BoldOblique"
                    elif font_name == "Times-Roman":
                        base_font = "Times-Roman"
                        bold_font = "Times-Bold"
                        italic_font = "Times-Italic"
                        bold_italic_font = "Times-BoldItalic"
                    else:  # Courier
                        base_font = "Courier"
                        bold_font = "Courier-Bold"
                        italic_font = "Courier-Oblique"
                        bold_italic_font = "Courier-BoldOblique"
                    
                    c.setFont(base_font, font_size)
                    
                    # Calculate position
                    line_height = font_size * 1.2
                    y_start = y * letter[1]
                    y_pos = letter[1] - y_start
                    
                    # First, combine all text while preserving formatting
                    combined_text = ""
                    format_ranges = []  # List of (start, end, formats) tuples
                    current_pos = 0
                    
                    for text, formats in processed_text:
                        if text:
                            start = current_pos
                            combined_text += text
                            end = current_pos + len(text)
                            if formats:  # Only store if there's formatting
                                format_ranges.append((start, end, formats))
                            current_pos = end
                    
                    # Wrap text according to max_chars
                    wrapped_lines = wrap_text(combined_text, max_chars)
                    
                    # Draw text line by line
                    x_offset = x * letter[0]
                    
                    for line in wrapped_lines:
                        if not line.strip():  # Handle empty lines
                            y_pos -= line_height
                            continue
                            
                        x_pos = x_offset
                        current_pos = combined_text.find(line)
                        
                        # Split line into segments based on formatting
                        segments = []
                        current_segment_start = 0
                        current_formats = []
                        
                        for i in range(len(line)):
                            pos_in_text = current_pos + i
                            new_formats = []
                            
                            # Find all formats that apply to this position
                            for start, end, formats in format_ranges:
                                if start <= pos_in_text < end:
                                    new_formats.extend(formats)
                            
                            # If formats changed, end current segment and start new one
                            if new_formats != current_formats:
                                if i > current_segment_start:
                                    segments.append((
                                        line[current_segment_start:i],
                                        current_formats
                                    ))
                                current_segment_start = i
                                current_formats = new_formats
                        
                        # Add final segment
                        if current_segment_start < len(line):
                            segments.append((
                                line[current_segment_start:],
                                current_formats
                            ))
                        
                        # Draw segments
                        for i, (segment, formats) in enumerate(segments):
                            # Apply font formatting
                            if 'italic' in formats:
                                current_font = italic_font
                                if 'bold' in formats:
                                    current_font = bold_italic_font
                            elif 'bold' in formats:
                                current_font = bold_font
                            else:
                                current_font = base_font
                            
                            c.setFont(current_font, font_size)
                            
                            # Draw text segment
                            c.drawString(x_pos, y_pos, segment)
                            
                            # Add underline if needed
                            if 'underline' in formats:
                                width = c.stringWidth(segment, current_font, font_size)
                                y_underline = y_pos - 1.5
                                c.setLineWidth(0.5)
                                c.line(x_pos, y_underline, x_pos + width, y_underline)
                            
                            x_pos += c.stringWidth(segment, current_font, font_size)
                        
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


def save_settings(settings):
    """Save settings to JSON file"""
    settings_file = "settings.json"
    try:
        with open(settings_file, 'w') as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        print(f"Error saving settings: {e}")

def load_settings():
    """Load settings from JSON file"""
    settings_file = "settings.json"
    default_settings = {
        "font_name": "Helvetica",
        "font_size": "12",
        "x_percent": "10",
        "y_percent": "20",
        "max_chars": "80",
        "filename_prefix": "emprius_{name}.pdf",
        "template_path": "",
        "csv_path": "",
        "output_dir": "",
        "text_content": ""
    }
    
    try:
        if os.path.exists(settings_file):
            with open(settings_file, 'r') as f:
                settings = json.load(f)
                # Update with any missing default settings
                for key, value in default_settings.items():
                    if key not in settings:
                        settings[key] = value
                return settings
    except Exception as e:
        print(f"Error loading settings: {e}")
    
    return default_settings

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

    # Load saved settings
    settings = load_settings()

    template_var = StringVar(value=settings["template_path"])
    csv_var = StringVar(value=settings["csv_path"])
    output_dir_var = StringVar(value=settings["output_dir"])
    filename_prefix_var = StringVar(value=settings["filename_prefix"])
    font_var = StringVar(value=settings["font_name"])
    font_size_var = StringVar(value=settings["font_size"])
    x_percent_var = StringVar(value=settings["x_percent"])
    y_percent_var = StringVar(value=settings["y_percent"])
    max_chars_var = StringVar(value=settings["max_chars"])

    # Function to save current settings
    def save_current_settings(*args):
        current_settings = {
            "template_path": template_var.get(),
            "csv_path": csv_var.get(),
            "output_dir": output_dir_var.get(),
            "filename_prefix": filename_prefix_var.get(),
            "font_name": font_var.get(),
            "font_size": font_size_var.get(),
            "x_percent": x_percent_var.get(),
            "y_percent": y_percent_var.get(),
            "max_chars": max_chars_var.get(),
            "text_content": text_widget.get("1.0", "end-1c") if text_widget.get("1.0", "end-1c") != "Type your text here..." else ""
        }
        save_settings(current_settings)

    # Track changes to save settings
    template_var.trace_add("write", save_current_settings)
    csv_var.trace_add("write", save_current_settings)
    output_dir_var.trace_add("write", save_current_settings)
    filename_prefix_var.trace_add("write", save_current_settings)
    font_var.trace_add("write", save_current_settings)
    font_size_var.trace_add("write", save_current_settings)
    x_percent_var.trace_add("write", save_current_settings)
    y_percent_var.trace_add("write", save_current_settings)
    max_chars_var.trace_add("write", save_current_settings)

    # Save settings when window is closed
    def on_closing():
        save_current_settings()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    headers_var = StringVar(value="Select a CSV file to see available tags.")

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

    Label(root, text="Text:").grid(row=3, column=0)
    
    # Add note about special characters
    Label(root, text="Use \\n to force new line, and \\t for tab", font=('TkDefaultFont', 8)).grid(row=3, column=1, sticky='w')
    
    # Create a frame for the toolbar
    toolbar_frame = Frame(root)
    toolbar_frame.grid(row=3, column=1, sticky='e')
    
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
    def apply_format(tag, font_config):
        try:
            # Get current selection
            try:
                selection_start = text_widget.index("sel.first")
                selection_end = text_widget.index("sel.last")
            except:
                messagebox.showinfo("Info", "Please select text to format")
                return
            
            # Check if tag already exists at this position
            current_tags = text_widget.tag_names(selection_start)
            
            if tag in current_tags:
                # Remove formatting
                text_widget.tag_remove(tag, selection_start, selection_end)
            else:
                # Add formatting
                text_widget.tag_add(tag, selection_start, selection_end)
                # Configure font for display
                if tag == "bold":
                    text_widget.tag_configure(tag, font=font.Font(weight="bold"))
                elif tag == "italic":
                    text_widget.tag_configure(tag, font=font.Font(slant="italic"))
                elif tag == "underline":
                    text_widget.tag_configure(tag, underline=True)
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def apply_bold():
        apply_format("bold", {"weight": "bold", "family": "TkDefaultFont"})
    
    def apply_italic():
        apply_format("italic", {"slant": "italic", "family": "TkDefaultFont"})
    
    def apply_underline():
        apply_format("underline", {"underline": True, "family": "TkDefaultFont"})
    
    # Create button styles
    style.configure('Bold.TButton', font=('TkDefaultFont', 10, 'bold'))
    style.configure('Italic.TButton', font=('TkDefaultFont', 10, 'italic'))
    style.configure('Underline.TButton', font=('TkDefaultFont', 10, 'underline'))
    
    # Create formatting buttons
    bold_button = TtkButton(toolbar_frame, text="B", width=3, command=apply_bold, style='Bold.TButton')
    italic_button = TtkButton(toolbar_frame, text="I", width=3, command=apply_italic, style='Italic.TButton')
    underline_button = TtkButton(toolbar_frame, text="U", width=3, command=apply_underline, style='Underline.TButton')
    
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
    
    text_input_help = "Type your text here. Use {tags} for CSV values. Tags are case-sensitive, e.g. {name} and are defined as the CSV column headers."

    # Clear default text when clicked
    def clear_default_text(event):
        if text_widget.get("1.0", "end-1c") == text_input_help:
            text_widget.delete("1.0", "end")
            text_widget.unbind('<Button-1>')
    
    # Load saved text content or show default
    saved_text = settings.get("text_content", "")
    if saved_text:
        text_widget.insert("1.0", saved_text)
    else:
        text_widget.insert("1.0", text_input_help)
        text_widget.bind('<Button-1>', clear_default_text)

    # Save text content when it changes
    def on_text_change(event=None):
        if text_widget.get("1.0", "end-1c") != text_input_help:
            save_current_settings()
    
    text_widget.bind('<<Modified>>', on_text_change)

    Label(root, textvariable=headers_var, wraplength=400, justify="left").grid(row=5, column=1)

    Label(root, text="Font:").grid(row=6, column=0)
    font_options = ["Helvetica", "Times-Roman", "Courier"]  # Built-in PDF fonts with style variants
    font_var.set("Helvetica")  # Default to Helvetica as it's always available
    OptionMenu(root, font_var, *font_options).grid(row=6, column=1)

    Label(root, text="Font Size:").grid(row=7, column=0)
    Entry(root, textvariable=font_size_var).grid(row=7, column=1)

    Label(root, text="Text X Position (%):").grid(row=8, column=0)
    Entry(root, textvariable=x_percent_var).grid(row=8, column=1)

    Label(root, text="Text Y Position (%):").grid(row=9, column=0)
    Entry(root, textvariable=y_percent_var).grid(row=9, column=1)

    Label(root, text="Chars per Line:").grid(row=10, column=0)
    Entry(root, textvariable=max_chars_var).grid(row=10, column=1)

    Label(root, text="Filename with tags:").grid(row=11, column=0)
    Entry(root, textvariable=filename_prefix_var).grid(row=11, column=1)

    Button(root, text="Generate PDFs", command=lambda: generate_pdfs(
        root,
        template_var.get(),
        csv_var.get(),
        output_dir_var.get(),
        filename_prefix_var.get(),
        get_custom_text(),
        font_var.get(),
        font_size_var.get(),
        x_percent_var.get(),
        y_percent_var.get(),
        max_chars_var.get(),
        text_widget
    )).grid(row=12, column=1)

    root.mainloop()


if __name__ == "__main__":
    main()
