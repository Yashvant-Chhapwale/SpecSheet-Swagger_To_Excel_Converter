import customtkinter as ctk
from urllib.parse import urlparse
from tkinter import filedialog
import tempfile
import sys
import os
import subprocess
import Logic


# === GUI Actions ===
current_mode = ctk.get_appearance_mode().lower()

def toggle_theme():
    current_theme = ctk.get_appearance_mode()
    theme = "Light" if current_theme == "Dark" else "Dark"
    ctk.set_appearance_mode(theme)
    current_mode = ctk.get_appearance_mode().lower()
    app.configure(fg_color=colors[current_mode]["primaryBg"])
    theme_button.configure(fg_color=colors[current_mode]["primaryBg"], hover_color=colors[current_mode]["primaryBg"])
    title.configure(text_color=colors[current_mode]["label"])
    entry.configure(fg_color=colors[current_mode]["entry"], border_color=colors[current_mode]["entry_border"], text_color=colors[current_mode]["entry_text"])
    convert_btn.configure(fg_color=colors[current_mode]["button"], hover_color=colors[current_mode]["button_hover"], text_color=colors[current_mode]["button_text"])
    result.configure(fg_color=colors[current_mode]["result"])
    result_btnFrame.configure(fg_color=colors[current_mode]["secondaryBg"])
    download_btn.configure(fg_color=colors[current_mode]["button_2"], hover_color=colors[current_mode]["button_hover_2"], text_color=colors[current_mode]["button_text_2"])
    preview_btn.configure(fg_color=colors[current_mode]["button_2"], hover_color=colors[current_mode]["button_hover_2"], text_color=colors[current_mode]["button_text_2"])
    theme_button.configure(text="◐" if theme == "Dark" else "☀")

def on_click_convert():
    global workbook
    swagger_url = entry.get().strip()
    result.pack_forget()
    alert.pack_forget()
    success_alert.pack_forget()
    result_btnFrame.pack_forget()

    if not swagger_url:
        alert.configure(text="⚠️ Please Enter a Valid URL!", text_color="#6A6A6A", fg_color="#FFC107")
        result.pack()
        alert.pack()
        return

    data = Logic.fetch_from_swagger(swagger_url)
    if not data:
        alert.configure(text="⚠️ Data Not Found!", text_color="#6A6A6A", fg_color="#FFC107")
        result.pack()
        alert.pack()
        return

    try:
        workbook = Logic.convert_to_excel(data)
        success_alert.configure(text="✨ Conversion Successfull!", text_color="#ffffff", fg_color="#0466c8")
        result.pack()
        success_alert.pack(pady=10)
        result_btnFrame.pack(fill="x")     
    except Exception as e:
        alert.configure(text=f"❌ 422: Unprocessable Entity", text_color="white", fg_color="#EB3434")
        result.pack()
        alert.pack()

def on_click_download():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        title="Export Spreadsheet As"
    )
    if file_path:
        try:
            workbook.save(file_path)
            success_alert.configure(text=f"📁 File Saved At: {file_path}", text_color="white", fg_color="#0466c8")
        except Exception as e:
            success_alert.configure(text=f"❌ 500: Save Unsuccessful", text_color="white", fg_color="#EB3434")

def on_click_preview():
    try:
        if workbook is None:
            success_alert.configure(text="⚠️ Workbook Not Available", text_color="#6A6A6A", fg_color="#FFC107")
            return

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            temp_path = tmp.name
            workbook.save(temp_path)

        if os.name == 'nt':  # Windows
            os.startfile(temp_path)
        elif os.name == 'posix':  # macOS / Linux
            subprocess.call(["open" if sys.platform == "darwin" else "xdg-open", temp_path])

        success_alert.configure(text="✨ Preview is Ready!", text_color="white", fg_color="#0466c8")
    except Exception as e:
        success_alert.configure(text=f"❌ Failed to Load Preview", text_color="white", fg_color="#EB3434")

# === GUI Setup ===
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue") 

app = ctk.CTk()
app.title("SpecSheet")
app.geometry("650x450")
app.minsize(650, 450)      

if getattr(sys, 'frozen', False): 
    base_path = sys._MEIPASS 
else: 
    base_path = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_path, "assets", "specSheetSymbol.ico") 
    logoLight = os.path.join(base_path, "assets", "specSheetLogoLight.png")
    logoDark = os.path.join(base_path, "assets", "specSheetLogoDark.png")
                                        
app.wm_iconbitmap(icon_path)

colors = {
    "light": {"primaryBg": "#eff6ff", "secondaryBg":"#86efac", "result":"#e5e5e5", "switch_text":"#4ade80", "label":"#4ade80", "entry":"#e5e5e5", "entry_text":"#008000", "entry_border":"#eff6ff", "button": "#22c55e", "button_hover": "#16a34a", "button_text":"#eff6ff", "button_2": "#eff6ff", "button_hover_2": "#dee2e6", "button_text_2":"#166534"},
    "dark": {"primaryBg": "#00132d", "secondaryBg":"#070c1c", "result":"#001e45", "switch_text":"#4ade80", "label":"#eff6ff","entry":"#070c1c",  "entry_text":"#eff6ff", "entry_border":"#335480", "button": "#4ade80", "button_hover": "#86efac", "button_text":"#16007a", "button_2": "#4ade80", "button_hover_2": "#86efac", "button_text_2":"#16007a"}
}

theme_button = ctk.CTkButton(app, text="◐" if current_mode == "dark" else "☀", font=("Verdana", 20, "bold"), command=toggle_theme, text_color=colors[current_mode]["switch_text"], fg_color=colors[current_mode]["primaryBg"],  hover_color=colors[current_mode]["primaryBg"], width=25, height=15, corner_radius=3)
theme_button.place(relx=0.95, rely=0.0, anchor="ne", x=1)

app.configure(fg_color=colors[current_mode]["primaryBg"])

title = ctk.CTkLabel(app, text="SpecSheet", font=("Verdana", 30, "normal"), text_color=colors[current_mode]["label"])
title.pack(pady=(80,40))

spacer = ctk.CTkLabel(app, text="", height=5)
spacer.pack()

entry = ctk.CTkEntry(app, placeholder_text="Enter Swagger/OpenAPI URL", font=("Verdana", 13, "normal"), width=400, height=35, text_color=colors[current_mode]["entry_text"], fg_color=colors[current_mode]["entry"], border_color=colors[current_mode]["entry_border"])
entry.pack(pady=5)

convert_btn = ctk.CTkButton(app, text="Generate SpecSheet", command=on_click_convert, font=("Verdana", 13, "normal"), corner_radius=3, text_color=colors[current_mode]["button_text"], fg_color=colors[current_mode]["button"], hover_color=colors[current_mode]["button_hover"],width=150, height=30)
convert_btn.pack(pady=20)

result= ctk.CTkFrame(app, width=400, height=200, corner_radius=5, fg_color=colors[current_mode]["result"])

alert = ctk.CTkLabel(result, text="", justify="center", corner_radius=5, width=400, height=35, font=("Verdana", 14, "normal"))
success_alert = ctk.CTkLabel(result, text="", justify="center", corner_radius=5, width=380, height=35, font=("Verdana", 14, "normal"))

result_btnFrame = ctk.CTkFrame(result, width=400, height=200, corner_radius=5, fg_color=colors[current_mode]["secondaryBg"])

download_btn = ctk.CTkButton(result_btnFrame, text="📥 Download xlsx", command=on_click_download, font=("Verdana", 13, "normal"), corner_radius=3, text_color=colors[current_mode]["button_text_2"], fg_color=colors[current_mode]["button_2"], hover_color=colors[current_mode]["button_hover_2"],width=150, height=30)
download_btn.grid(row=0, column=0, rowspan=1, columnspan=1, sticky="nsew", padx=20, pady=20, ipadx=10, ipady=5)

preview_btn = ctk.CTkButton(result_btnFrame, text="🔍 Preview", command=on_click_preview, font=("Verdana", 13, "normal"), corner_radius=3, text_color=colors[current_mode]["button_text_2"], fg_color=colors[current_mode]["button_2"], hover_color=colors[current_mode]["button_hover_2"], width=150, height=30)
preview_btn.grid(row=0, column=1, rowspan=1, columnspan=1, sticky="nsew", padx=20, pady=20, ipadx=10, ipady=5)

app.mainloop()
