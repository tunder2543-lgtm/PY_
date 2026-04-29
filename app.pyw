import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
import sys
import json
import glob
from datetime import datetime
import textwrap

def get_app_dir():
    """คืนค่า directory ที่ไฟล์ .exe หรือ script ตั้งอยู่"""
    if getattr(sys, 'frozen', False):
        # ถ้ารันเป็น .exe (PyInstaller)
        return os.path.dirname(sys.executable)
    else:
        # ถ้ารันเป็น script ปกติ
        return os.path.dirname(os.path.abspath(__file__))

def get_resource_path(relative_path):
    """คืนค่า path ของ resource ที่ถูกรวมเข้าไปใน PyInstaller bundle"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# =====================================================================
# ค่าคงที่เริ่มต้น (Global Settings)
# =====================================================================
TEXT_COORDS = {
    'Item':    [1180, 415],
    'SKU':     [1176, 509],
    'Gam':     [1176, 574],
    'Color':   [1176, 632],
    'Setting': [1178, 694],
    'Date':    [1176, 752],
}

FONT_SIZES = {
    'Item':    30,
    'SKU':     30,
    'Gam':     30,
    'Color':   30,
    'Setting': 30,
    'Date':    30,
}

FONT_BOLDS = {
    'Item':    False,
    'SKU':     False,
    'Gam':     False,
    'Color':   False,
    'Setting': False,
    'Date':    False,
}

FONT_NAME = "Chonburi"

OPTIONS = {
    'Color': ["", "Black", "White", "Green", "Dark Green", "Red", "Yellow", "Blue", "Gray", "Colorful", "Purple"],
    'Setting': ["", "Rhodium", "CZ", "Stainless", "925silver", "Micron"],
    'Gam': ["", "Natural Jade Jadeite / (A-Jade)", "Natural Jade Nephrite"]
}

def get_font_for_image(size, is_bold=False):
    """โหลดฟอนต์ตามขนาดที่กำหนด สำหรับบันทึกภาพจริง"""
    local_font_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Fonts')
    sys_font_dir = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
    app_dir = get_app_dir()
    
    font_paths = [
        os.path.join(app_dir, "Chonburi-Regular.ttf"), # check app dir first
        os.path.join(app_dir, "Chonburi.ttf"),
        "Chonburi-Regular.ttf", # directory ปัจจุบัน
        "Chonburi.ttf",
        os.path.join(local_font_dir, "Chonburi-Regular.ttf"),
        os.path.join(local_font_dir, "Chonburi.ttf"),
        os.path.join(sys_font_dir, "Chonburi-Regular.ttf"),
        os.path.join(sys_font_dir, "Chonburi.ttf")
    ]
    
    fallback_paths = [
        os.path.join(sys_font_dir, "tahomabd.ttf" if is_bold else "tahoma.ttf"),
        os.path.join(sys_font_dir, "arialbd.ttf" if is_bold else "arial.ttf"),
        "tahoma.ttf",
        "arial.ttf"
    ]
    
    for path in font_paths + fallback_paths:
        try:
            return ImageFont.truetype(path, size)
        except IOError:
            continue
    return ImageFont.load_default()

def find_product_image(sku, folder_path):
    """ค้นหารูปภาพสินค้าจาก SKU ในโฟลเดอร์ที่กำหนด
    รองรับนามสกุล .jpg, .jpeg, .png, .webp, .bmp
    ค้นหาทั้งตรงชื่อ และแบบ contains (SKU เป็นส่วนหนึ่งของชื่อไฟล์)
    """
    if not folder_path or not os.path.isdir(folder_path):
        return None
    
    extensions = ['*.jpg', '*.jpeg', '*.png', '*.webp', '*.bmp']
    safe_sku = sku.strip()
    
    # 1) ค้นหาแบบตรงชื่อ (exact match โดยไม่สนนามสกุล)
    for ext in extensions:
        pattern = os.path.join(folder_path, ext)
        for filepath in glob.glob(pattern):
            basename = os.path.splitext(os.path.basename(filepath))[0]
            if basename.lower() == safe_sku.lower():
                return filepath
    
    # 2) ค้นหาแบบ contains (SKU เป็นส่วนหนึ่งของชื่อไฟล์)
    for ext in extensions:
        pattern = os.path.join(folder_path, ext)
        for filepath in glob.glob(pattern):
            basename = os.path.splitext(os.path.basename(filepath))[0]
            if safe_sku.lower() in basename.lower():
                return filepath
    
    return None


class PreviewEditor(ctk.CTkToplevel):
    def __init__(self, master, template_path, output_dir, records, product_image_dir=None):
        super().__init__(master)
        
        self.title("I REAL JADE - ตัวอย่างและการแก้ไข (Preview & Editor)")
        self.geometry("1400x850")
        
        self.template_path = template_path
        self.output_dir = output_dir
        self.records = records
        self.current_idx = 0
        self.product_image_dir = product_image_dir  # โฟลเดอร์รูปภาพสินค้า
        self.product_tk_image = None  # เก็บ reference รูปภาพสินค้า
        self._product_img_cache = {}  # cache รูปภาพตาม SKU ป้องกัน garbage collection
        
        # เตรียมข้อมูล per-record (ตำแหน่ง/ขนาด/Bold อิสระแต่ละหน้า)
        for rec in self.records:
            if '_coords' not in rec:
                rec['_coords'] = {f: list(v) for f, v in TEXT_COORDS.items()}
            if '_sizes' not in rec:
                rec['_sizes'] = dict(FONT_SIZES)
            if '_bolds' not in rec:
                rec['_bolds'] = dict(FONT_BOLDS)
        
        self.pil_image = Image.open(self.template_path).convert("RGB")
        self.img_w, self.img_h = self.pil_image.size
        
        self.scale_factor = 1.0
        self.offset_x = 0
        self.offset_y = 0
        self.selected_field = None
        self.drag_mode = None
        
        self.create_widgets()
        self.load_record(0)
        self.grab_set()

    def create_widgets(self):
        # ---------------- แถบบน (Navigation / บันทึก) ----------------
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=10)
        
        btn_prev = ctk.CTkButton(top_frame, text="◀ ก่อนหน้า", font=("Chonburi", 16), command=self.prev_record)
        btn_prev.pack(side="left", padx=10, pady=10)
        
        self.lbl_counter = ctk.CTkLabel(top_frame, text="0 / 0", font=("Chonburi", 18, "bold"))
        self.lbl_counter.pack(side="left", padx=20)
        
        btn_next = ctk.CTkButton(top_frame, text="ถัดไป ▶", font=("Chonburi", 16), command=self.next_record)
        btn_next.pack(side="left", padx=10)
        
        btn_save_all = ctk.CTkButton(top_frame, text="💾 ยืนยันและบันทึกรูปทั้งหมด", font=("Chonburi", 18, "bold"), 
                                     fg_color="#1F6032", hover_color="#174825", height=45, command=self.save_all)
        btn_save_all.pack(side="right", padx=10)
        
        # ---------------- โซนหลัก (ซ้าย Canvas, ขวา Controls) ----------------
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=(0,10))
        
        # --- ซ้าย: Canvas ---
        self.canvas_container = ctk.CTkFrame(main_frame)
        self.canvas_container.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        self.canvas = tk.Canvas(self.canvas_container, bg="#333333", bd=0, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        
        # เหตุการณ์ bindings
        self.canvas.bind("<ButtonPress-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_drag_motion)
        self.canvas.bind("<ButtonRelease-1>", self.on_drag_stop)
        
        # --- ขวา: Editor ---
        right_frame = ctk.CTkFrame(main_frame, width=420)
        right_frame.pack(side="right", fill="y")
        right_frame.pack_propagate(False) 
        
        # --- แพนลรูปภาพสินค้า (Product Image Panel) ---
        if self.product_image_dir:
            self.product_img_frame = ctk.CTkFrame(right_frame, fg_color="#1a1a2e", corner_radius=12)
            self.product_img_frame.pack(fill="x", padx=8, pady=(8, 4))
            
            # Header ของแพนลรูปภาพ
            img_header = ctk.CTkFrame(self.product_img_frame, fg_color="transparent")
            img_header.pack(fill="x", padx=8, pady=(6, 2))
            
            ctk.CTkLabel(img_header, text="\U0001f4f7 รูปภาพสินค้า", 
                         font=("Chonburi", 14, "bold"), text_color="#00d4aa").pack(side="left")
            
            self.product_sku_label = ctk.CTkLabel(img_header, text="", 
                                                   font=("Tahoma", 11), text_color="#888888")
            self.product_sku_label.pack(side="right")
            
            # พื้นที่แสดงรูป
            self.product_img_container = ctk.CTkFrame(self.product_img_frame, fg_color="#0d0d1a", 
                                                       corner_radius=8, height=200)
            self.product_img_container.pack(fill="x", padx=8, pady=(2, 8))
            self.product_img_container.pack_propagate(False)
            
            self.product_img_label = ctk.CTkLabel(self.product_img_container, text="", 
                                                   image=None, fg_color="transparent")
            self.product_img_label.pack(expand=True)
            
            # Label สถานะ
            self.product_status_label = ctk.CTkLabel(self.product_img_frame, text="กำลังโหลด...", 
                                                      font=("Tahoma", 10), text_color="#666666")
            self.product_status_label.pack(pady=(0, 6))
        
        ctk.CTkLabel(right_frame, text="\U0001f4dd เครื่องมือแก้ไข (Editor)", font=("Chonburi", 20, "bold"), text_color="#1F6032").pack(pady=(10,0))
        ctk.CTkLabel(right_frame, text="คลิกข้อความในรูปเพื่อโชว์กรอบ ย้าย หรือย่อขยายตรงมุม!", text_color="gray", font=("Tahoma", 12)).pack(pady=(0,8))
        
        scrollable_controls = ctk.CTkScrollableFrame(right_frame)
        scrollable_controls.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.field_vars = {}
        self.size_vars = {}
        self.bold_vars = {}
        self.cb_vars = {}
        self.pos_vars = {}
        
        for field in TEXT_COORDS.keys():
            box = ctk.CTkFrame(scrollable_controls)
            box.pack(fill="x", padx=5, pady=5)
            
            # Header
            header_frame = ctk.CTkFrame(box, fg_color="transparent")
            header_frame.pack(fill="x", padx=5, pady=(5,0))
            
            ctk.CTkLabel(header_frame, text=f"{field}", font=("Chonburi", 14, "bold")).pack(side="left")
            
            pos_var = tk.StringVar(value=f"X: {TEXT_COORDS[field][0]}, Y: {TEXT_COORDS[field][1]}")
            self.pos_vars[field] = pos_var
            ctk.CTkLabel(header_frame, textvariable=pos_var, font=("Tahoma", 11, "bold"), text_color="#00A8FF").pack(side="left", padx=15)
            
            # ตัวหนา (Bold)
            bold_var = tk.BooleanVar(value=FONT_BOLDS[field])
            self.bold_vars[field] = bold_var
            cb_bold = ctk.CTkCheckBox(header_frame, text="ตัวหนา", variable=bold_var, command=lambda f=field: self.toggle_bold(f), width=60)
            cb_bold.pack(side="right", padx=5)
            
            # ขนาด (Size)
            size_var = tk.IntVar(value=FONT_SIZES[field])
            self.size_vars[field] = size_var
            
            btn_add = ctk.CTkButton(header_frame, text="+", width=25, height=25, command=lambda f=field: self.change_size(f, 2))
            btn_add.pack(side="right")
            ctk.CTkLabel(header_frame, textvariable=size_var, width=25).pack(side="right", padx=2)
            btn_sub = ctk.CTkButton(header_frame, text="-", width=25, height=25, command=lambda f=field: self.change_size(f, -2))
            btn_sub.pack(side="right", padx=(5, 0))
            ctk.CTkLabel(header_frame, text="ขนาด:", font=("Tahoma", 12)).pack(side="right")
            
            # ช่องกรอกแก้ไข
            if field == 'Item':
                tb = ctk.CTkTextbox(box, font=("Tahoma", 14), height=50)
                tb.pack(fill="x", padx=10, pady=(2, 2))
                tb.bind("<KeyRelease>", lambda e, f=field, t=tb: self.on_textbox_change(f, t))
                self.field_vars[field] = tb
            elif field in OPTIONS:
                # สำหรับ Gam ให้ใช้ Textbox เผื่อขึ้นบรรทัดใหม่
                if field == 'Gam':
                    tb = ctk.CTkTextbox(box, font=("Tahoma", 14), height=50)
                    tb.pack(fill="x", padx=10, pady=(2, 2))
                    tb.bind("<KeyRelease>", lambda e, f=field, t=tb: self.on_textbox_change(f, t))
                    self.field_vars[field] = tb
                else:
                    var = tk.StringVar()
                    var.trace_add("write", lambda name, index, mode, f=field, v=var: self.on_entry_change(f, v))
                    en = ctk.CTkEntry(box, textvariable=var, font=("Tahoma", 14), height=30)
                    en.pack(fill="x", padx=10, pady=(2, 2))
                    self.field_vars[field] = var
                    
                # สร้าง Checkboxes
                cb_frame = ctk.CTkFrame(box, fg_color="transparent")
                cb_frame.pack(fill="x", padx=10, pady=2)
                
                self.cb_vars[field] = {}
                row_idx, col_idx = 0, 0
                max_cols = 1 if field == 'Gam' else 3
                
                for opt in OPTIONS[field]:
                    if not opt: continue
                    cb_var = tk.BooleanVar(value=False)
                    self.cb_vars[field][opt] = cb_var
                    
                    cb = ctk.CTkCheckBox(cb_frame, text=opt, variable=cb_var, width=10,
                                         command=lambda f=field: self.on_checkbox_click(f))
                    cb.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="w")
                    col_idx += 1
                    if col_idx >= max_cols:
                        col_idx = 0
                        row_idx += 1
                        
            else:
                var = tk.StringVar()
                var.trace_add("write", lambda name, index, mode, f=field, v=var: self.on_entry_change(f, v))
                en = ctk.CTkEntry(box, textvariable=var, font=("Tahoma", 14), height=30)
                en.pack(fill="x", padx=10, pady=(2, 2))
                self.field_vars[field] = var

    # --- Control Behaviours ---
    def change_size(self, field, delta):
        record = self.records[self.current_idx]
        new_size = max(10, self.size_vars[field].get() + delta)
        self.size_vars[field].set(new_size)
        record['_sizes'][field] = new_size
        self.redraw_canvas()
        
    def toggle_bold(self, field):
        record = self.records[self.current_idx]
        record['_bolds'][field] = self.bold_vars[field].get()
        self.redraw_canvas()

    def on_entry_change(self, field, stringvar):
        self.records[self.current_idx][field] = stringvar.get()
        self.redraw_canvas()
        
    def on_textbox_change(self, field, textbox):
        new_val = textbox.get("1.0", "end-1c")
        self.records[self.current_idx][field] = new_val
        self.redraw_canvas()
        
    def on_checkbox_click(self, field):
        selected = [opt for opt, var in self.cb_vars[field].items() if var.get()]
        sep = "\n" if field == 'Gam' else ", "
        new_text = sep.join(selected)
        
        self.records[self.current_idx][field] = new_text
        
        widget = self.field_vars[field]
        if isinstance(widget, ctk.CTkTextbox):
            widget.delete("1.0", "end")
            widget.insert("1.0", new_text)
        else:
            widget.set(new_text)
            
        self.redraw_canvas()
        
    # --- Canvas Behaviours ---
    def load_record(self, idx):
        self.current_idx = idx
        record = self.records[idx]
        self.lbl_counter.configure(text=f"ใบที่ {idx+1} / {len(self.records)}")
        
        self.selected_field = None # รีเซ็ต selection เมื่อเปลี่ยนใบ
        
        for field, var_or_tb in self.field_vars.items():
            val = record.get(field, "")
            if isinstance(var_or_tb, ctk.CTkTextbox):
                var_or_tb.delete("1.0", "end")
                var_or_tb.insert("1.0", val)
            else:
                var_or_tb.set(val)
                
            # ซิงค์ Checkboxes ให้ตรงกับข้อความที่มีอยู่
            if field in self.cb_vars:
                for opt, cb_var in self.cb_vars[field].items():
                    if opt in val:
                        cb_var.set(True)
                    else:
                        cb_var.set(False)
        
        # ซิงค์ขนาด/Bold/ตำแหน่ง ตาม record นี้
        for field in TEXT_COORDS.keys():
            self.size_vars[field].set(record['_sizes'][field])
            self.bold_vars[field].set(record['_bolds'][field])
            cx, cy = record['_coords'][field]
            self.pos_vars[field].set(f"X: {cx}, Y: {cy}")
        
        # โหลดรูปภาพสินค้า (ถ้าเปิดโหมดนี้)
        self.load_product_image(record.get('SKU', ''))
                        
        self.redraw_canvas()
    
    def _clear_product_image(self):
        """เคลียร์รูปภาพสินค้าออกจาก Label ให้แสดงมืดเปล่า (แก้ปัญหา CTkLabel ไม่ยอมเคลียร์รูปเก่า)"""
        # สร้างรูปเปล่า 1x1 pixel สีดำเพื่อบังคับให้ CTkLabel ลบรูปเก่าออก
        if not hasattr(self, '_blank_image'):
            blank = Image.new("RGB", (1, 1), color=(13, 13, 26))  # สีเดียวกับ bg ของ container
            self._blank_image = ImageTk.PhotoImage(blank)
        self.product_tk_image = self._blank_image
        self.product_img_label.configure(image=self._blank_image, text="")

    def load_product_image(self, sku):
        """โหลดและแสดงรูปภาพสินค้าตาม SKU (มี cache ป้องกันรูปหายเมื่อเปลี่ยนหน้า)"""
        if not self.product_image_dir or not hasattr(self, 'product_img_label'):
            return
        
        sku = str(sku).strip()
        if not sku or sku == 'nan':
            self._clear_product_image()
            self.product_img_label.configure(text="ไม่มีรหัส SKU")
            self.product_sku_label.configure(text="")
            self.product_status_label.configure(text="ไม่พบ SKU", text_color="#ff6b6b")
            return
        
        self.product_sku_label.configure(text=f"SKU: {sku}")
        
        # ตรวจสอบ cache ก่อน - ถ้ามีรูปแล้วไม่ต้องโหลดใหม่
        if sku in self._product_img_cache:
            cached = self._product_img_cache[sku]
            self.product_tk_image = cached['image']
            self.product_img_label.configure(
                image=self.product_tk_image, 
                text=""
            )
            self.product_status_label.configure(
                text=f"\u2714 {cached['filename']}", 
                text_color="#00d4aa"
            )
            return
        
        # ค้นหารูปภาพ
        img_path = find_product_image(sku, self.product_image_dir)
        
        if img_path:
            try:
                pil_img = Image.open(img_path).convert("RGB")
                
                # ปรับขนาดให้พอดีกับแพนล (max 380x180)
                max_w, max_h = 380, 180
                orig_w, orig_h = pil_img.size
                scale = min(max_w / orig_w, max_h / orig_h)
                new_w = int(orig_w * scale)
                new_h = int(orig_h * scale)
                pil_img = pil_img.resize((new_w, new_h), Image.LANCZOS)
                
                self.product_tk_image = ImageTk.PhotoImage(pil_img)
                
                # เก็บ cache เพื่อป้องกัน garbage collection
                filename = os.path.basename(img_path)
                self._product_img_cache[sku] = {
                    'image': self.product_tk_image,
                    'filename': filename
                }
                
                self.product_img_label.configure(
                    image=self.product_tk_image, 
                    text=""
                )
                
                self.product_status_label.configure(
                    text=f"\u2714 {filename}", 
                    text_color="#00d4aa"
                )
            except Exception as e:
                self._clear_product_image()
                self.product_img_label.configure(text="ไม่สามารถเปิดรูปได้")
                self.product_status_label.configure(
                    text=f"Error: {str(e)[:40]}", 
                    text_color="#ff6b6b"
                )
        else:
            self._clear_product_image()
            self.product_img_label.configure(
                text=f"ไม่พบรูปภาพสำหรับ SKU: {sku}"
            )
            self.product_status_label.configure(
                text="\u2716 ไม่พบรูปภาพในโฟลเดอร์", 
                text_color="#ff6b6b"
            )

    def on_canvas_resize(self, event):
        cw = event.width
        ch = event.height
        if cw < 10 or ch < 10: return
        
        ratio_w = cw / self.img_w
        ratio_h = ch / self.img_h
        self.scale_factor = min(ratio_w, ratio_h)
        
        new_w = int(self.img_w * self.scale_factor)
        new_h = int(self.img_h * self.scale_factor)
        
        # Resize image for display
        display_image = self.pil_image.resize((new_w, new_h), Image.LANCZOS)
        self.tk_image = ImageTk.PhotoImage(display_image)
        
        # Center image mathematically
        self.offset_x = (cw - new_w) // 2
        self.offset_y = (ch - new_h) // 2
        
        self.redraw_canvas()

    def redraw_canvas(self):
        self.canvas.delete("all")
        if not hasattr(self, 'tk_image'): return
        
        # วาดรูป Background ขึ้นมาก่อน
        self.canvas.create_image(self.offset_x, self.offset_y, anchor="nw", image=self.tk_image, tags="background")
        
        record = self.records[self.current_idx]
        rec_coords = record['_coords']
        rec_sizes = record['_sizes']
        rec_bolds = record['_bolds']
        
        for field in TEXT_COORDS.keys():
            orig_x, orig_y = rec_coords[field]
            text = record.get(field, "")
            orig_size = rec_sizes[field]
            is_bold = rec_bolds[field]
            
            # คำนวณพิกัดและสเกลตามหน้าจอ
            disp_x = self.offset_x + int(orig_x * self.scale_factor)
            disp_y = self.offset_y + int(orig_y * self.scale_factor)
            disp_size = max(1, int(orig_size * self.scale_factor)) # ป้องกัน size เป็น 0
            
            weight = "bold" if is_bold else "normal"
            try:
                draw_font = (FONT_NAME, -disp_size, weight)
                item_id = self.canvas.create_text(disp_x, disp_y, text=text, font=draw_font, anchor="nw", fill="black", tags=("text_item", field))
            except:
                draw_font = ("Tahoma", -disp_size, weight)
                item_id = self.canvas.create_text(disp_x, disp_y, text=text, font=draw_font, anchor="nw", fill="black", tags=("text_item", field))
                
            # วาดกรอบ selection
            if self.selected_field == field:
                bbox = self.canvas.bbox(item_id)
                if bbox:
                    pad = 4
                    x1, y1, x2, y2 = bbox[0]-pad, bbox[1]-pad, bbox[2]+pad, bbox[3]+pad
                    self.canvas.create_rectangle(x1, y1, x2, y2, outline="#00A8FF", width=2, dash=(4,4), tags="selection_box")
                    # จุด handle สำหรับย่อขยายด้านมุมขวาล่าง
                    hx1, hy1 = x2-6, y2-6
                    hx2, hy2 = x2+6, y2+6
                    self.canvas.create_rectangle(hx1, hy1, hx2, hy2, fill="#00A8FF", outline="white", tags=("handle", field))

    def on_canvas_click(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        
        # เช็คว่าคลิกโดนอะไรไหม
        items = self.canvas.find_withtag("current")
        if not items:
            self.selected_field = None
            self.redraw_canvas()
            return
            
        tags = self.canvas.gettags(items[0])
        
        record = self.records[self.current_idx]
        
        # ถ้าคลิกโดนปุ่ม Handle ตรงมุม
        if "handle" in tags:
            self.drag_mode = "resize"
            self.selected_field = tags[1]
            self.drag_start_y = y
            self.start_size = record['_sizes'][self.selected_field]
            return
            
        # ถ้าคลิกโดนข้อความธรรมดา
        if "text_item" in tags:
            self.drag_mode = "move"
            self.selected_field = next((t for t in tags if t in TEXT_COORDS), None)
            self.drag_start_x = x
            self.drag_start_y = y
            
            self.redraw_canvas() # redraw วาดกรอบครอบ
            
            # เมื่อ redraw เสร็จ item id ตัวเก่าจะถูกลบไป ต้องหา id ของข้อความนี้ใหม่
            new_items = self.canvas.find_withtag(self.selected_field)
            if new_items:
                self.drag_item = new_items[0]
                
            self.canvas.config(cursor="fleur")
            return

    def on_drag_motion(self, event):
        if not self.drag_mode: return
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        
        record = self.records[self.current_idx]
        
        if self.drag_mode == "move" and hasattr(self, "drag_item"):
            dx = x - self.drag_start_x
            dy = y - self.drag_start_y
            self.canvas.move(self.drag_item, dx, dy)
            self.canvas.move("selection_box", dx, dy)
            self.canvas.move("handle", dx, dy)
            self.drag_start_x = x
            self.drag_start_y = y
            
            # อัปเดตพิกัดแบบ Live
            if self.selected_field:
                coords = self.canvas.coords(self.drag_item)
                ox = int((coords[0] - self.offset_x) / self.scale_factor)
                oy = int((coords[1] - self.offset_y) / self.scale_factor)
                self.pos_vars[self.selected_field].set(f"X: {ox}, Y: {oy}")
            
        elif self.drag_mode == "resize" and getattr(self, "selected_field", None):
            dy = y - self.drag_start_y
            delta_size = int((dy / 3) / self.scale_factor) # ชดเชยการอัพสเกลของจอ
            new_size = max(10, self.start_size + delta_size)
            if new_size != record['_sizes'][self.selected_field]:
                record['_sizes'][self.selected_field] = new_size
                self.size_vars[self.selected_field].set(new_size)
                self.redraw_canvas()

    def on_drag_stop(self, event):
        self.canvas.config(cursor="")
        if self.drag_mode == "move" and self.selected_field:
            record = self.records[self.current_idx]
            coords = self.canvas.coords(self.drag_item)
            # ต้องสเกลพิกัดกลับไปเป็นไซส์รูปต้นฉบับจริง
            orig_x = int((coords[0] - self.offset_x) / self.scale_factor)
            orig_y = int((coords[1] - self.offset_y) / self.scale_factor)
            record['_coords'][self.selected_field] = [orig_x, orig_y]
            self.pos_vars[self.selected_field].set(f"X: {orig_x}, Y: {orig_y}")
        self.drag_mode = None

    def prev_record(self):
        if self.current_idx > 0: self.load_record(self.current_idx - 1)
            
    def next_record(self):
        if self.current_idx < len(self.records) - 1: self.load_record(self.current_idx + 1)
            
    # --- Execute Image Generation ---
    def save_all(self):
        success_count = 0
        total = len(self.records)
        self.config(cursor="wait"); self.update()
        
        try:
            for record in self.records:
                sku_val = str(record['SKU']).strip()
                if not sku_val or sku_val == 'nan': continue
                
                img = Image.open(self.template_path).convert("RGB")
                draw = ImageDraw.Draw(img)
                
                rec_coords = record['_coords']
                rec_sizes = record['_sizes']
                rec_bolds = record['_bolds']
                
                for field in TEXT_COORDS.keys():
                    coords = rec_coords[field]
                    text_val = str(record.get(field, ""))
                    size = rec_sizes[field]
                    is_bold = rec_bolds[field]
                    
                    if text_val and text_val != 'nan':
                        f = get_font_for_image(size, is_bold)
                        stroke_width = 1 if is_bold else 0 # ทำขอบหนา 1 pixel เพื่อให้ดูเป็นตัวอักษรหนา
                        draw.text(coords, text_val, font=f, fill="black", stroke_width=stroke_width, stroke_fill="black")
                        
                safe_sku = "".join(c for c in sku_val if c.isalnum() or c in "-_ ")
                output_file_path = os.path.join(self.output_dir, f"{safe_sku}.jpg")
                img.save(output_file_path, "JPEG", quality=95)
                success_count += 1
                
            self.config(cursor="")
            messagebox.showinfo("สำเร็จ!", f"บันทึกรูปภาพอัตโนมัติเรียบร้อย!\nจำนวน {success_count} จากทั้งหมด {total} ใบ")
            self.destroy()
        except Exception as e:
            self.config(cursor="")
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการบันทึก:\n{str(e)}")


# =====================================================================
# Main Application
# =====================================================================
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("green")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("I REAL JADE - Automation Configurator")
        self.geometry("750x830")
        self.resizable(False, False)
        
        self.ui_font = ("Chonburi", 16)
        
        self.template_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.product_img_enabled = tk.BooleanVar(value=False)
        self.product_img_dir = tk.StringVar()
        
        self.create_widgets()
        self.load_config()
        
    def load_config(self):
        config_path = os.path.join(get_app_dir(), "config.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.template_path.set(data.get('template', ''))
                    self.excel_path.set(data.get('excel', ''))
                    self.output_dir.set(data.get('out_dir', ''))
                    self.product_img_enabled.set(data.get('product_img_enabled', False))
                    self.product_img_dir.set(data.get('product_img_dir', ''))
                    # อัปเดต UI ตาม state ที่โหลดมา
                    self.toggle_product_image_mode()
            except:
                pass
                
    def save_config(self):
        data = {
            'template': self.template_path.get(),
            'excel': self.excel_path.get(),
            'out_dir': self.output_dir.get(),
            'product_img_enabled': self.product_img_enabled.get(),
            'product_img_dir': self.product_img_dir.get()
        }
        try:
            config_path = os.path.join(get_app_dir(), "config.json")
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False)
        except:
            pass
        
    def create_widgets(self):
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(pady=20, padx=20, fill="x")
        ctk.CTkLabel(header_frame, text="✨ ระบบดึงข้อมูลสำหรับสร้างใบรับประกัน", font=("Chonburi", 28, "bold"), text_color="#1F6032").pack()
        
        platform_frame = ctk.CTkFrame(self)
        platform_frame.pack(pady=5, padx=20, fill="x")
        ctk.CTkLabel(platform_frame, text="เลือก Platform ข้อมูล:", font=self.ui_font).pack(side="left", padx=10, pady=10)
        self.platform_var = ctk.StringVar(value="TIKTOK")
        ctk.CTkSegmentedButton(platform_frame, values=["TIKTOK", "FACEBOOK"], variable=self.platform_var, font=("Chonburi", 14)).pack(side="left", padx=10)
        
        file_frame = ctk.CTkFrame(self)
        file_frame.pack(pady=10, padx=20, fill="x")
        self.create_file_row(file_frame, "\U0001f5bc\ufe0f รูป Template:", self.template_path, self.browse_template)
        self.create_file_row(file_frame, "\U0001f4ca ไฟล์ Excel:", self.excel_path, self.browse_excel)
        self.create_file_row(file_frame, "\U0001f4c1 โฟลเดอร์เก็บรูป:", self.output_dir, self.browse_output)
        
        # ==================== ส่วนเสริม: โหมดรูปภาพสินค้า ====================
        product_img_frame = ctk.CTkFrame(self, border_width=2, border_color="#2196F3", corner_radius=10)
        product_img_frame.pack(pady=(5, 5), padx=20, fill="x")
        
        # แถวบน: Checkbox เปิด/ปิด
        product_header = ctk.CTkFrame(product_img_frame, fg_color="transparent")
        product_header.pack(fill="x", padx=10, pady=(8, 2))
        
        self.cb_product_img = ctk.CTkCheckBox(
            product_header, 
            text="\U0001f4f7 เปิดโหมดแสดงรูปภาพสินค้า (ดึงจาก SKU)", 
            font=("Chonburi", 14, "bold"),
            variable=self.product_img_enabled,
            command=self.toggle_product_image_mode,
            text_color="#1565C0"
        )
        self.cb_product_img.pack(side="left")
        
        # แถวล่าง: เลือกโฟลเดอร์
        self.product_img_row = ctk.CTkFrame(product_img_frame, fg_color="transparent")
        self.product_img_row.pack(fill="x", padx=10, pady=(2, 8))
        
        ctk.CTkLabel(self.product_img_row, text="   โฟลเดอร์รูปสินค้า:", 
                     font=("Tahoma", 13), text_color="#555").pack(side="left")
        self.product_img_entry = ctk.CTkEntry(self.product_img_row, textvariable=self.product_img_dir, 
                                               width=330, font=("Tahoma", 12), state="disabled")
        self.product_img_entry.pack(side="left", padx=5)
        self.btn_browse_product = ctk.CTkButton(self.product_img_row, text="ค้นหา..", width=80, 
                                                 font=("Chonburi", 12), command=self.browse_product_img)
        self.btn_browse_product.pack(side="left", padx=5)
        
        # ซ่อนแถวเลือกโฟลเดอร์ไว้ก่อน (จะแสดงเมื่อติ้กเลือก)
        self.product_img_row.pack_forget()
        
        # กล่องตั้งค่า Dropdown
        data_frame = ctk.CTkFrame(self)
        data_frame.pack(pady=10, padx=20, fill="both", expand=True)
        ctk.CTkLabel(data_frame, text="\U0001f4dd ตั้งค่าข้อมูลที่จะถูกใส่เป็นค่า Default", font=("Chonburi", 18, "bold")).grid(row=0, column=0, columnspan=2, pady=10, sticky="w", padx=10)
        
        # Color
        ctk.CTkLabel(data_frame, text="COLOR (สี): *", font=self.ui_font).grid(row=1, column=0, padx=20, pady=10, sticky="e")
        self.color_var = ctk.StringVar(value=OPTIONS['Color'][0])
        self.color_dropdown = ctk.CTkComboBox(data_frame, values=OPTIONS['Color'], variable=self.color_var, font=self.ui_font, width=250)
        self.color_dropdown.grid(row=1, column=1, sticky="w")
        
        # Setting
        ctk.CTkLabel(data_frame, text="SETTING:", font=self.ui_font).grid(row=2, column=0, padx=20, pady=10, sticky="e")
        self.setting_var = ctk.StringVar(value=OPTIONS['Setting'][0])
        self.setting_dropdown = ctk.CTkComboBox(data_frame, values=OPTIONS['Setting'], variable=self.setting_var, font=self.ui_font, width=250)
        self.setting_dropdown.grid(row=2, column=1, sticky="w")
        
        # Gem
        ctk.CTkLabel(data_frame, text="GEM (อัญมณี):", font=self.ui_font).grid(row=3, column=0, padx=20, pady=10, sticky="e")
        self.gem_var = ctk.StringVar(value=OPTIONS['Gam'][0])
        self.gem_dropdown = ctk.CTkComboBox(data_frame, values=OPTIONS['Gam'], variable=self.gem_var, font=self.ui_font, width=250)
        self.gem_dropdown.grid(row=3, column=1, sticky="w")
        
        # Date Default เป็น "(เว้นไว้)" ไว้ก่อนเผื่อไม่อยากกำหนด
        ctk.CTkLabel(data_frame, text="DATE (วันที่): *", font=self.ui_font).grid(row=4, column=0, padx=20, pady=10, sticky="e")
        self.date_entry = ctk.CTkEntry(data_frame, font=self.ui_font, width=150)
        self.date_entry.insert(0, "") # เว้นไว้ก่อนได้
        self.date_entry.grid(row=4, column=1, sticky="w")
        ctk.CTkButton(data_frame, text="ใช้วันที่วันนี้", width=80, command=lambda: [self.date_entry.delete(0,'end'), self.date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))]).grid(row=4, column=1, padx=160, sticky="w")

        btn_run = ctk.CTkButton(self, text="\u26a1 ดึงข้อมูลและเปิดหน้า Preview Editor", font=("Chonburi", 18, "bold"), 
                                fg_color="#1F6032", hover_color="#174825", height=50, corner_radius=10,
                                command=self.load_data)
        btn_run.pack(pady=20, padx=20, fill="x")

    def create_file_row(self, parent, label_text, str_var, browse_cmd):
        row_frame = ctk.CTkFrame(parent, fg_color="transparent")
        row_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(row_frame, text=label_text, font=self.ui_font, width=120, anchor="e").pack(side="left", padx=10)
        entry = ctk.CTkEntry(row_frame, textvariable=str_var, width=400, font=("Tahoma", 12), state="disabled")
        entry.pack(side="left", padx=5)
        ctk.CTkButton(row_frame, text="ค้นหา..", width=80, font=("Chonburi", 12), command=browse_cmd).pack(side="left", padx=5)

    def browse_template(self):
        filename = filedialog.askopenfilename(title="เลือกไฟล์ Template", filetypes=[("Image files", "*.jpg;*.jpeg;*.png")])
        if filename: 
            self.template_path.set(filename)
            self.save_config()
            
    def browse_excel(self):
        filename = filedialog.askopenfilename(title="เลือกไฟล์ข้อมูล Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filename: 
            self.excel_path.set(filename)
            self.save_config()
            
    def browse_output(self):
        dirname = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับเก็บผลลัพธ์")
        if dirname: 
            self.output_dir.set(dirname)
            self.save_config()
    
    def toggle_product_image_mode(self):
        """เปิด/ปิดโหมดแสดงรูปภาพสินค้า"""
        if self.product_img_enabled.get():
            self.product_img_row.pack(fill="x", padx=10, pady=(2, 8))
        else:
            self.product_img_row.pack_forget()
        self.save_config()
    
    def browse_product_img(self):
        dirname = filedialog.askdirectory(title="เลือกโฟลเดอร์ที่เก็บรูปภาพสินค้า (ชื่อไฟล์ = SKU)")
        if dirname:
            self.product_img_dir.set(dirname)
            self.save_config()

    def load_data(self):
        template = self.template_path.get()
        excel_file = self.excel_path.get()
        out_dir = self.output_dir.get()
        
        if not template or not excel_file or not out_dir:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกไฟล์ให้ครบทั้ง 3 ช่องครับ")
            return
            
        color_val = self.color_var.get()
        date_val = self.date_entry.get().strip()
        setting_val = self.setting_var.get()
        gem_val = self.gem_var.get()
        
        try:
            platform = self.platform_var.get()
            
            if platform == "TIKTOK":
                df = pd.read_excel(excel_file)
                
                if 'Seller SKU' not in df.columns or 'Product Name' not in df.columns:
                    messagebox.showerror("Error", "ไม่พบคอลัมน์ 'Seller SKU' หรือ 'Product Name'")
                    return
                    
                if len(df) > 0 and str(df['Seller SKU'].iloc[0]).startswith("Seller sku Input"):
                    df = df.iloc[1:]
                    
                df = df.dropna(subset=['Seller SKU'])
                df['Seller SKU'] = df['Seller SKU'].astype(str).str.strip()
                df_unique = df.drop_duplicates(subset=['Seller SKU'])
                
                if len(df_unique) == 0:
                    messagebox.showinfo("แจ้งเตือน", "ไม่พบข้อมูลรหัสสินค้าในไฟล์ Excel!")
                    return
                    
                records = []
                
                for index, row in df_unique.iterrows():
                    sku_val = str(row['Seller SKU']).strip()
                    item_name = str(row['Product Name']).strip()
                    if sku_val == 'nan' or not sku_val: continue
                    
                    # ตัดข้อความ Item (ให้ขึ้นบรรทัดใหม่ทุกๆ 40 ตัวอักษร โดยไม่ตัดคำทิ้ง)
                    wrapped = textwrap.wrap(item_name, width=40, break_long_words=True)
                    item_display = "\n".join(wrapped)

                        
                    records.append({
                        'Item': item_display,
                        'SKU': sku_val,
                        'Gam': gem_val,
                        'Color': color_val,
                        'Setting': setting_val,
                        'Date': date_val
                    })
                    
            elif platform == "FACEBOOK":
                xls = pd.ExcelFile(excel_file)
                sheet_names = xls.sheet_names
                
                # ค้นหา Sheet2 ด้วยวิธีที่ครอบคลุมมากขึ้น
                target_sheet = None
                
                # 1) ลองค้นหาจากชื่อ sheet ที่มีคำว่า sheet2 หรือ sheets2
                for sn in sheet_names:
                    sn_clean = sn.replace(" ", "").replace("_", "").upper()
                    if sn_clean in ["SHEETS2", "SHEET2", "SHEETS_2", "SHEET_2"]:
                        target_sheet = sn
                        break
                
                # 2) ลองใช้ชีตที่ 2 (index 1) ถ้ามีมากกว่า 1 ชีต
                if target_sheet is None and len(sheet_names) >= 2:
                    target_sheet = sheet_names[1]
                
                # 3) ถ้ามีชีตเดียวก็ใช้ชีตแรก
                if target_sheet is None:
                    target_sheet = sheet_names[0]
                
                df = pd.read_excel(excel_file, sheet_name=target_sheet, header=None)
                
                if len(df.columns) < 3:
                    messagebox.showerror("Error", f"ไฟล์ข้อมูลมีคอลัมน์ไม่พอในชีต '{target_sheet}' (ต้องการอย่างน้อย 3 คอลัมน์ A, B, C)")
                    return
                    
                df = df.dropna(subset=[1])
                df[1] = df[1].astype(str).str.strip()
                df_unique = df.drop_duplicates(subset=[1])
                
                if len(df_unique) == 0:
                    messagebox.showinfo("แจ้งเตือน", f"ไม่พบข้อมูลรหัสสินค้าในไฟล์ Excel ชีต '{target_sheet}'!")
                    return
                    
                records = []
                for index, row in df_unique.iterrows():
                    col_a = str(row[0]).strip() if pd.notna(row[0]) and str(row[0]) != 'nan' else ""
                    col_b = str(row[1]).strip() if pd.notna(row[1]) and str(row[1]) != 'nan' else ""
                    col_c = str(row[2]).strip() if pd.notna(row[2]) and str(row[2]) != 'nan' else ""
                    
                    if not col_b or col_b.lower() in ['sku', 'รหัสสต็อก', 'รหัสสต๊อก', 'none', 'nan']: 
                        continue
                        
                    sku_val = col_b
                    item_text = f"{col_a} {col_c}".strip()
                    
                    wrapped = textwrap.wrap(item_text, width=40, break_long_words=True)
                    item_display = "\n".join(wrapped)
                        
                    records.append({
                        'Item': item_display,
                        'SKU': sku_val,
                        'Gam': gem_val,
                        'Color': color_val,
                        'Setting': setting_val,
                        'Date': date_val
                    })
                
            # ส่ง product_image_dir ไปยัง PreviewEditor ถ้าเปิดโหมดนี้
            product_dir = None
            if self.product_img_enabled.get():
                product_dir = self.product_img_dir.get()
                if not product_dir or not os.path.isdir(product_dir):
                    messagebox.showwarning("คำเตือน", "เปิดโหมดรูปภาพสินค้าอยู่ แต่ยังไม่ได้เลือกโฟลเดอร์รูปภาพ\nจะเปิดหน้า Preview โดยไม่แสดงรูปสินค้า")
                    product_dir = None
            
            PreviewEditor(self, template, out_dir, records, product_image_dir=product_dir)
            
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถโหลดข้อมูลได้:\n{str(e)}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
