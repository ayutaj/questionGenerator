import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import config
import Converter
from pathlib import Path
from helper import to_camel_case  # Ensure helper.to_camel_case exists


DOWNLOADS = str(Path.cwd() / "NewOutput")


# -----------------------------
# Load config defaults (for UI only)
# -----------------------------
def load_word_config():
    # Note: keys here must match the keys used in apply_word_settings()
    return {
        "Topic Name": config.TOPIC_NAME,
        "Font (Question)": config.QUESTION_FONT,
        "Font Size (Question)": config.QUESTION_FONT_SIZE,
        "Font (Option)": config.OPTION_FONT,
        "Font Size (Option)": config.OPTION_FONT_SIZE,
        "Option Spacing (tab)": getattr(config, "tab", 3),
        "Left Offset (tab2)": getattr(config, "tab2", 5),
        "Margins (top,bottom,left,right)": (
            config.TOP_MARGIN_CM,
            config.BOTTOM_MARGIN_CM,
            config.LEFT_MARGIN_CM,
            config.RIGHT_MARGIN_CM
        )
    }


def load_ppt_config():
    return {
        "Topic Name": config.TOPIC_NAME,
        "PPT Font (Question)": config.PPT_QUESTION_FONT,
        "Question Size": config.PPT_QUESTION_FONT_SIZE,
        "PPT Font (Option)": config.PPT_OPTION_FONT,
        "Option Size": config.PPT_OPTION_FONT_SIZE
    }


# -----------------------------
# Apply Word/PPT settings from UI
# -----------------------------
def apply_word_settings(entries, output_folder):
    # Topic Name affects everything
    config.TOPIC_NAME = entries["Topic Name"].get()
    topic_cc = to_camel_case(config.TOPIC_NAME)

    # Update output files dynamically
    config.outputfoldername = topic_cc
    config.OUTPUT_FOLDER = os.path.join(output_folder, topic_cc)
    config.OUTPUT_QUESTIONS_FILE = os.path.join(config.OUTPUT_FOLDER, f"{topic_cc}_questions.docx")
    config.OUTPUT_ANSWERS_FILE = os.path.join(config.OUTPUT_FOLDER, f"{topic_cc}_answer_key.docx")
    config.OUTPUT_PPT_FILE = os.path.join(config.OUTPUT_FOLDER, f"{topic_cc}_questions.pptx")

    os.makedirs(config.OUTPUT_FOLDER, exist_ok=True)

    # Other fields
    config.QUESTION_FONT = entries["Font (Question)"].get()
    config.QUESTION_FONT_SIZE = int(entries["Font Size (Question)"].get())
    config.OPTION_FONT = entries["Font (Option)"].get()
    config.OPTION_FONT_SIZE = int(entries["Font Size (Option)"].get())

    # NEW - tab spacing (Option Spacing and Left Offset)
    # Use int conversion and fallback in case user leaves blank or invalid input
    try:
        config.tab = int(entries["Option Spacing (tab)"].get())
    except Exception:
        config.tab = getattr(config, "tab", 3)

    try:
        config.tab2 = int(entries["Left Offset (tab2)"].get())
    except Exception:
        config.tab2 = getattr(config, "tab2", 5)

    # Margins
    margins_val = entries["Margins (top,bottom,left,right)"].get()
    try:
        margins = [m.strip() for m in margins_val.split(",")]
        config.TOP_MARGIN_CM = float(margins[0])
        config.BOTTOM_MARGIN_CM = float(margins[1])
        config.LEFT_MARGIN_CM = float(margins[2])
        config.RIGHT_MARGIN_CM = float(margins[3])
    except Exception:
        # keep existing values on parse error
        pass


def apply_ppt_settings(entries, output_folder):
    # Topic name also applies for PPT filename
    config.TOPIC_NAME = entries["Topic Name"].get()
    topic_cc = to_camel_case(config.TOPIC_NAME)

    config.outputfoldername = topic_cc
    config.OUTPUT_FOLDER = os.path.join(output_folder, topic_cc)
    config.OUTPUT_PPT_FILE = os.path.join(config.OUTPUT_FOLDER, f"{topic_cc}_questions.pptx")

    os.makedirs(config.OUTPUT_FOLDER, exist_ok=True)

    try:
        config.PPT_QUESTION_FONT = entries["PPT Font (Question)"].get()
        config.PPT_QUESTION_FONT_SIZE = int(entries["Question Size"].get())
        config.PPT_OPTION_FONT = entries["PPT Font (Option)"].get()
        config.PPT_OPTION_FONT_SIZE = int(entries["Option Size"].get())
    except Exception:
        pass


# -----------------------------
# MAIN SCREEN
# -----------------------------
root = tk.Tk()
root.title("Question Generator")
root.geometry("480x360")
root.resizable(False, False)

tk.Label(
    root,
    text="Question Paper Generator",
    font=("Arial", 16, "bold")
).pack(pady=20)


# -----------------------------
# FORM WINDOW
# -----------------------------
def open_form(mode):
    root.withdraw()

    form = tk.Toplevel(root)
    form.title(f"{mode} Settings")
    form.geometry("620x700")
    form.resizable(False, False)

    def go_home():
        form.destroy()
        root.deiconify()

    # Header
    header = tk.Frame(form)
    header.pack(fill="x", pady=10)

    tk.Label(header, text=f"{mode} Settings",
             font=("Arial", 16, "bold")).pack(side="left", padx=20)

    tk.Button(header, text="🏠 Home", font=("Arial", 12),
              command=go_home).pack(side="right", padx=20)

    settings = load_word_config() if mode == "Word" else load_ppt_config()
    entries = {}

    body = tk.Frame(form)
    body.pack(pady=10)

    # Create form fields
    for i, (label, default) in enumerate(settings.items()):
        tk.Label(body, text=label, font=("Arial", 10, "bold")).grid(
            row=i, column=0, sticky="w", pady=6, padx=10
        )

        if isinstance(default, tuple):
            default = ", ".join(str(x) for x in default)

        entry = tk.Entry(body, width=50)
        entry.insert(0, str(default))
        entry.grid(row=i, column=1, pady=6, padx=(0, 10))
        entries[label] = entry

    # -----------------------------
    # Excel selector
    # -----------------------------
    row_offset = len(settings)

    tk.Label(body, text="Select Excel File:", font=("Arial", 10, "bold")).grid(
        row=row_offset, column=0, pady=10, padx=10
    )

    excel_entry = tk.Entry(body, width=50)
    excel_entry.grid(row=row_offset, column=1, padx=(0, 10))

    def browse_excel():
        file = filedialog.askopenfilename(
            title="Choose Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if file:
            excel_entry.delete(0, tk.END)
            excel_entry.insert(0, file)

    tk.Button(body, text="Browse", command=browse_excel).grid(
        row=row_offset, column=2, padx=5
    )

    # -----------------------------
    # OUTPUT FOLDER SELECTOR
    # -----------------------------
    row_of = row_offset + 1

    tk.Label(body, text="Select Output Folder:", font=("Arial", 10, "bold")).grid(
        row=row_of, column=0, pady=10, padx=10
    )

    output_entry = tk.Entry(body, width=50)
    output_entry.insert(0, DOWNLOADS)  # default
    output_entry.grid(row=row_of, column=1, padx=(0, 10))

    def browse_output():
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            output_entry.delete(0, tk.END)
            output_entry.insert(0, folder)

    tk.Button(body, text="Browse", command=browse_output).grid(
        row=row_of, column=2, padx=5
    )

    # -----------------------------
    # GENERATE
    # -----------------------------
    def generate():
        excel_file = excel_entry.get().strip()
        output_folder = output_entry.get().strip()

        if not excel_file:
            messagebox.showerror("Error", "Please select an Excel file.")
            return

        os.makedirs(config.INPUT_FOLDER, exist_ok=True)

        # Copy only if source != dest
        if os.path.abspath(excel_file) != os.path.abspath(config.INPUT_FILE):
            shutil.copy(excel_file, config.INPUT_FILE)

        df = Converter.read_questions_from_excel(config.INPUT_FILE)

        # Apply settings based on mode (apply settings before generating)
        if mode == "Word":
            apply_word_settings(entries, output_folder)
            Converter.write_questions_to_word(df, config.OUTPUT_QUESTIONS_FILE)
            Converter.write_answer_key_to_word(df, config.OUTPUT_ANSWERS_FILE)
        else:
            apply_ppt_settings(entries, output_folder)
            Converter.write_ppt(df, config.OUTPUT_PPT_FILE)

        messagebox.showinfo("Success", f"{mode} files generated!\nSaved in: {config.OUTPUT_FOLDER}")
        try:
            os.startfile(config.OUTPUT_FOLDER)
        except Exception:
            pass

    tk.Button(
        form, text="Generate",
        command=generate,
        font=("Arial", 12, "bold"),
        bg="green", fg="white"
    ).pack(pady=20)


# -----------------------------
# MAIN SCREEN BUTTONS
# -----------------------------
tk.Button(
    root,
    text="Generate Word Document",
    font=("Arial", 12),
    width=36,
    command=lambda: open_form("Word")
).pack(pady=10)

tk.Button(
    root,
    text="Generate PPT",
    font=("Arial", 12),
    width=36,
    command=lambda: open_form("PPT")
).pack(pady=10)

root.mainloop()
