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
            
        # Skip other whitespace at start of line
        if not current_line and word.isspace():
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
        # Count and process all tabs in the line
        while '\\t' in line:
            # Find position of tab
            tab_pos = line.find('\\t')
            # Add spaces at tab position
            line = line[:tab_pos] + '    ' + line[tab_pos+2:]
        processed_lines.append(line)
    
    # Join lines back together
    text = '\n'.join(processed_lines)
    
    # Process other escape sequences
    text = text.replace('\\r', '\r')
    return text


def generate_pdfs(template_path, csv_path, output_dir, filename_prefix, custom_text, font_name, font_size, x_percent, y_percent, max_chars, text_widget):
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
                    # Get text and formatting
                    text = text_widget.get("1.0", "end-1c")
                    
                    # Get text with formatting by segments
                    formatted_text = []
                    start = "1.0"
                    while True:
                        if text_widget.compare(start, ">=", "end-1c"):
                            break
                            
                        # Get current formatting
                        formats = []
                        for tag in ["bold", "italic", "underline"]:
                            if tag in text_widget.tag_names(start):
                                formats.append(tag)
                        
                        # Find next format change or space
                        next_pos = start
                        while True:
                            if text_widget.compare(next_pos, ">=", "end-1c"):
                                break
                            
                            next_char = text_widget.get(next_pos)
                            next_formats = []
                            for tag in ["bold", "italic", "underline"]:
                                if tag in text_widget.tag_names(next_pos):
                                    next_formats.append(tag)
                                    
                            # Break if formatting changes or we hit a space
                            if next_formats != formats or next_char.isspace():
                                break
                                
                            next_pos = text_widget.index(f"{next_pos}+1c")
                        
                        # Get text segment
                        text = text_widget.get(start, next_pos)
                        if text:
                            # Add segment with its formatting
                            formatted_text.append((text, formats))
                            
                        # If we stopped at a space, add it as a separate segment
                        if text_widget.compare(next_pos, "<", "end-1c"):
                            space_char = text_widget.get(next_pos)
                            if space_char.isspace():
                                # Add space as a separate segment without formatting
                                formatted_text.append((space_char, []))
                                next_pos = text_widget.index(f"{next_pos}+1c")
                            
                        start = next_pos
                    
                    # Process text with tags and escape sequences
                    processed_text = []
                    for text, formats in formatted_text:
                        # Replace tags
                        for key, value in row.items():
                            placeholder = "{" + key + "}"
                            text = text.replace(placeholder, str(value))
                        
                        # Process escape sequences
                        text = process_escape_sequences(text)
                        
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
                    
                    # Draw text with continuous flow
                    x_offset = x * letter[0]
                    current_line = []  # List of (text, format) tuples for current line
                    
                    for text, formats in processed_text:
                        if not text.strip():
                            # Draw current line if exists
                            if current_line:
                                x_pos = x_offset
                                for segment, seg_formats in current_line:
                                    # Apply font formatting
                                    # Apply font formatting with proper italic handling
                                    if 'italic' in seg_formats:
                                        current_font = italic_font
                                        if 'bold' in seg_formats:
                                            current_font = bold_italic_font
                                    elif 'bold' in seg_formats:
                                        current_font = bold_font
                                    else:
                                        current_font = base_font
                                    
                                    # Set font and draw text segment
                                    c.setFont(current_font, font_size)
                                    c.drawString(x_pos, y_pos, segment)
                                    
                                    # Add underline if needed
                                    if 'underline' in seg_formats:
                                        width = c.stringWidth(segment, current_font, font_size)
                                        y_underline = y_pos - 1.5
                                        c.setLineWidth(0.5)
                                        c.line(x_pos, y_underline, x_pos + width, y_underline)
                                    
                                    x_pos += c.stringWidth(segment, current_font, font_size)
                                y_pos -= line_height
                                current_line = []
                            y_pos -= line_height
                            continue
                        
                        # Split text into words
                        words = text.split()
                        for word in words:
                            # Calculate width of current line plus new word
                            test_width = x_offset
                            for segment, seg_formats in current_line:
                                # Apply font formatting with proper italic handling
                                if 'italic' in seg_formats:
                                    current_font = italic_font
                                    if 'bold' in seg_formats:
                                        current_font = bold_italic_font
                                elif 'bold' in seg_formats:
                                    current_font = bold_font
                                else:
                                    current_font = base_font
                                test_width += c.stringWidth(segment, current_font, font_size)
                            
                            # Calculate width of new word
                            if 'italic' in formats:
                                current_font = italic_font
                                if 'bold' in formats:
                                    current_font = bold_italic_font
                            elif 'bold' in formats:
                                current_font = bold_font
                            else:
                                current_font = base_font
                            
                            # Add space width if not first word on line
                            space_width = c.stringWidth(" ", current_font, font_size) if current_line else 0
                            word_width = c.stringWidth(word, current_font, font_size)
                            total_width = test_width + space_width + word_width
                            
                            # Check if word fits on current line
                            if total_width > letter[0] * 0.8:  # 80% of page width
                                # Draw current line
                                if current_line:
                                    x_pos = x_offset
                                    line_text = ""
                                    line_segments = []
                                    
                                    # First collect all segments
                                    for segment, seg_formats in current_line:
                                        if segment.strip():  # Only add non-whitespace segments
                                            line_segments.append((segment, seg_formats))
                                            line_text += segment
                                    
                                    # Now draw the collected segments
                                    for i, (segment, seg_formats) in enumerate(line_segments):
                                        # Apply font formatting
                                        if 'italic' in seg_formats:
                                            current_font = italic_font
                                            if 'bold' in seg_formats:
                                                current_font = bold_italic_font
                                        elif 'bold' in seg_formats:
                                            current_font = bold_font
                                        else:
                                            current_font = base_font
                                        
                                        c.setFont(current_font, font_size)
                                        
                                        # Add space between words, but not after the last word
                                        if i > 0:
                                            x_pos += c.stringWidth(" ", current_font, font_size)
                                        
                                        c.drawString(x_pos, y_pos, segment.strip())
                                        
                                        if 'underline' in seg_formats:
                                            width = c.stringWidth(segment.strip(), current_font, font_size)
                                            y_underline = y_pos - 1.5
                                            c.setLineWidth(0.5)
                                            c.line(x_pos, y_underline, x_pos + width, y_underline)
                                        
                                        x_pos += c.stringWidth(segment.strip(), current_font, font_size)
                                    
                                    y_pos -= line_height
                                    current_line = []
                                
                                # Start new line with current word
                                current_line.append((word, formats))
                            else:
                                # Add word to current line
                                if current_line:
                                    current_line.append((" ", []))  # Add unformatted space
                                current_line.append((word, formats))
                    
                    # Draw any remaining text in current line
                    if current_line:
                        x_pos = x_offset
                        line_segments = []
                        
                        # First collect all segments
                        for segment, seg_formats in current_line:
                            if segment.strip():  # Only add non-whitespace segments
                                line_segments.append((segment, seg_formats))
                        
                        # Now draw the collected segments
                        for i, (segment, seg_formats) in enumerate(line_segments):
                            # Apply font formatting
                            if 'italic' in seg_formats:
                                current_font = italic_font
                                if 'bold' in seg_formats:
                                    current_font = bold_italic_font
                            elif 'bold' in seg_formats:
                                current_font = bold_font
                            else:
                                current_font = base_font
                            
                            c.setFont(current_font, font_size)
                            
                            # Add space between words, but not after the last word
                            if i > 0:
                                x_pos += c.stringWidth(" ", current_font, font_size)
                            
                            c.drawString(x_pos, y_pos, segment.strip())
                            
                            if 'underline' in seg_formats:
                                width = c.stringWidth(segment.strip(), current_font, font_size)
                                y_underline = y_pos - 1.5
                                c.setLineWidth(0.5)
                                c.line(x_pos, y_underline, x_pos + width, y_underline)
                            
                            x_pos += c.stringWidth(segment.strip(), current_font, font_size)
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
    filename_prefix_var = StringVar(value="e.g. emprius_{name}.pdf")
    font_var = StringVar()
    font_size_var = StringVar(value="12")
    x_percent_var = StringVar(value="10")
    y_percent_var = StringVar(value="20")
    max_chars_var = StringVar(value="80")
    headers_var = StringVar(value="You can use \\n for new line and \\t for tab (4 spaces). Select a CSV file to see available tags.")

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
    Label(root, text="Use \\n for new line, \\t for tab (4 spaces)", font=('TkDefaultFont', 8)).grid(row=3, column=1, sticky='w')
    
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
    
    # Clear default text when clicked
    def clear_default_text(event):
        if text_widget.get("1.0", "end-1c") == "Type your text here...":
            text_widget.delete("1.0", "end")
            text_widget.unbind('<Button-1>')
    
    text_widget.insert("1.0", "Type your text here...")
    text_widget.bind('<Button-1>', clear_default_text)

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
