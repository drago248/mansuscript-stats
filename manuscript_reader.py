import sys
import re
import time
import tkinter as tk
from tkinter import filedialog, simpledialog, scrolledtext

def count_words(text):
    return len(text.strip().split())

def count_sentences(text):
    sentences = re.split(r'[.!?]+', text)
    return len([s for s in sentences if s.strip()])

def count_characters(text, with_spaces=True):
    return len(text) if with_spaces else len(text.replace(" ", ""))

def format_duration(minutes_float):
    total_seconds = int(round(minutes_float * 60))
    hours = total_seconds // 3600
    remainder = total_seconds % 3600
    mins = remainder // 60
    secs = remainder % 60
    if hours > 0:
        return f"{hours}h {mins}m {secs}s"
    elif mins > 0:
        return f"{mins}m {secs}s"
    else:
        return f"{secs}s"

def format_compute_time(elapsed_s):
    if elapsed_s < 0.1:
        ms = elapsed_s * 1000
        return f"{ms:.1f} ms"
    elif elapsed_s < 60:
        return f"{elapsed_s:.2f} s"
    else:
        mins = int(elapsed_s // 60)
        secs = int(elapsed_s % 60)
        if mins < 60:
            return f"{mins} min {secs} s"
        else:
            hours = mins // 60
            mins_remainder = mins % 60
            return f"{hours} h {mins_remainder} min {secs} s"

def extract_text_docx(filepath):
    try:
        import docx
    except ImportError:
        raise RuntimeError("python-docx library is required to read DOCX files.")
    doc = docx.Document(filepath)
    texts = [para.text for para in doc.paragraphs if para.text.strip()]
    return "\n".join(texts)

def extract_text_pdf(filepath):
    try:
        import PyPDF2
    except ImportError:
        raise RuntimeError("PyPDF2 library is required to read PDF files.")
    texts = []
    with open(filepath, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                texts.append(text)
    return "\n".join(texts)

def extract_text_txt(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        return f.read()

def extract_text_rtf(filepath):
    try:
        from striprtf.striprtf import rtf_to_text
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            rtf_content = f.read()
        return rtf_to_text(rtf_content)
    except ImportError:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        text = re.sub(r'{\\.*?}|\\[A-Za-z]+\d* ?', '', content)
        text = re.sub(r'[{}]', '', text)
        text = re.sub(r'\n+', '\n', text)
        return text

def extract_text_odt(filepath):
    try:
        from odf.opendocument import load as odf_load
        from odf.text import P as odf_p
    except ImportError:
        raise RuntimeError("odfpy library is required to read ODT files.")
    doc = odf_load(filepath)
    allparas = doc.getElementsByType(odf_p)
    texts = []
    for p in allparas:
        txt = []
        for node in p.childNodes:
            if node.nodeType == node.TEXT_NODE:
                txt.append(node.data)
        texts.append("".join(txt))
    return "\n".join(texts)

def extract_text(filepath):
    ext = filepath.lower().rsplit('.', 1)[-1]
    print(f"Detected file extension: {ext}")  # Debug print
    if ext == "docx":
        return extract_text_docx(filepath)
    elif ext == "pdf":
        return extract_text_pdf(filepath)
    elif ext == "txt":
        return extract_text_txt(filepath)
    elif ext == "rtf":
        return extract_text_rtf(filepath)
    elif ext == "odt":
        return extract_text_odt(filepath)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

def analyze_text(text, wpm_min=200, wpm_max=280):
    start_time = time.perf_counter()

    paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
    total_paragraphs = len(paragraphs)
    total_words = sum(count_words(p) for p in paragraphs)
    total_sentences = sum(count_sentences(p) for p in paragraphs)
    total_chars_with_spaces = len(text)
    total_chars_no_spaces = len(text.replace(" ", ""))

    results = []
    results.append(f"Total word count: {total_words}")
    results.append(f"Total paragraph count (non-empty): {total_paragraphs}")
    results.append(f"Sentence count: {total_sentences}")
    results.append(f"Character count (with spaces): {total_chars_with_spaces}")
    results.append(f"Character count (without spaces): {total_chars_no_spaces}")
    results.append(f"Estimated reading times based on WPM range [{wpm_min} - {wpm_max}] progressing in steps of 10:")

    for wpm in range(wpm_min, wpm_max + 1, 10):
        est_time_minutes = total_words / wpm if wpm != 0 else 0
        est_time_str = format_duration(est_time_minutes)
        results.append(f"  @ {wpm} WPM: {est_time_str}")

    end_time = time.perf_counter()
    elapsed = end_time - start_time
    results.append(f"\n[Compute time: {format_compute_time(elapsed)}]")
    return "\n".join(results)

def analyze_file(filepath, wpm_min=200, wpm_max=280):
    try:
        text = extract_text(filepath)
    except Exception as e:
        return f"Error reading file: {e}"
    return analyze_text(text, wpm_min, wpm_max)

class ManuscriptAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Manuscript Analyzer")
        self.is_dark = False

        self.frame = tk.Frame(root, padx=10, pady=10)
        self.frame.pack(fill=tk.BOTH, expand=True)

        btn_frame = tk.Frame(self.frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_select = tk.Button(btn_frame, text="Select File", command=self.on_select_file)
        self.btn_select.pack(side=tk.LEFT)

        self.btn_toggle_theme = tk.Button(btn_frame, text="Switch to Dark Mode", command=self.toggle_theme)
        self.btn_toggle_theme.pack(side=tk.RIGHT)

        self.text_area = scrolledtext.ScrolledText(self.frame, width=80, height=30, state=tk.DISABLED, wrap=tk.WORD)
        self.text_area.pack(fill=tk.BOTH, expand=True)

        self.apply_theme()

    def apply_theme(self):
        if self.is_dark:
            bg = '#1e1e1e'
            fg = '#d4d4d4'
            btn_bg = '#3c3f41'
            btn_fg = '#ffffff'
            self.root.configure(bg=bg)
            self.frame.configure(bg=bg)
            self.text_area.configure(bg=bg, fg=fg, insertbackground=fg)
            self.btn_select.configure(bg=btn_bg, fg=btn_fg, activebackground='#505050')
            self.btn_toggle_theme.configure(bg=btn_bg, fg=btn_fg, activebackground='#505050')
        else:
            bg = '#ffffff'
            fg = '#000000'
            btn_bg = '#f0f0f0'
            btn_fg = '#000000'
            self.root.configure(bg=bg)
            self.frame.configure(bg=bg)
            self.text_area.configure(bg=bg, fg=fg, insertbackground=fg)
            self.btn_select.configure(bg=btn_bg, fg=btn_fg, activebackground='#d9d9d9')
            self.btn_toggle_theme.configure(bg=btn_bg, fg=btn_fg, activebackground='#d9d9d9')

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self.apply_theme()
        self.btn_toggle_theme.config(text="Switch to Light Mode" if self.is_dark else "Switch to Dark Mode")

    def on_select_file(self):
        filepath = filedialog.askopenfilename(
            title="Select Document",
            filetypes=[
                ("Word Documents", "*.docx"),
                ("PDF Files", "*.pdf"),
                ("Text Files", "*.txt"),
                ("Rich Text Format", "*.rtf"),
                ("OpenDocument Text", "*.odt"),
                ("All Files", "*.*"),
            ]
        )
        if not filepath:
            return

        ext = filepath.lower().rsplit('.', 1)[-1]
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, f"Selected file: {filepath}\nDetected extension: {ext}\n\nProcessing...\n")
        self.text_area.config(state=tk.DISABLED)

        wpm_min = simpledialog.askinteger("Input", "Enter minimum WPM (default 200):", initialvalue=200, minvalue=1)
        if wpm_min is None:
            return
        wpm_max = simpledialog.askinteger("Input", "Enter maximum WPM (default 280):", initialvalue=280, minvalue=wpm_min)
        if wpm_max is None:
            return

        result_text = analyze_file(filepath, wpm_min, wpm_max)

        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, result_text)
        self.text_area.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = ManuscriptAnalyzerApp(root)
    root.mainloop()
