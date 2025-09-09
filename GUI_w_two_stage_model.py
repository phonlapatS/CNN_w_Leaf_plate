# GUI_two_stage.py
# -*- coding: utf-8 -*-
# ติดตั้งที่ต้องใช้ (ถ้ายังไม่มี):
# pip install customtkinter ultralytics opencv-python pillow firebase-admin numpy

import customtkinter as ctk
import tkinter as tk
from tkinter import font as tkfont
from tkinter import filedialog, messagebox

import cv2
import numpy as np
from PIL import Image, ImageTk
from ultralytics import YOLO

from datetime import datetime, timezone
import os, sys, json, csv, time, uuid, signal
import urllib.request, urllib.error

# -------- Firebase Admin SDK --------
import firebase_admin
from firebase_admin import credentials, db


class LeafPlateTwoStageApp:
    """
    Two-Stage Inference:
      Stage-1: shape_model  -> circle/heart/rectangle
      Stage-2: defect_model -> crack/hole
    - รวมผล → แสดง/บันทึก/ส่ง Firebase
    - ป้องกันปัญหา Ctrl+C ด้วย safe_after + SIGINT hook
    """

    # -----------------------------
    # Layout constants for Full HD
    # -----------------------------
    def set_layout_constants(self):
        self.W, self.H = 1920, 1080
        self.M = 25
        self.HEADER_Y, self.HEADER_H = 20, 90
        self.TOP_Y = 120
        self.TOP_H = 640
        self.GAP = self.M
        self.header_w = self.W - 2 * self.M
        self.left_w = int(self.header_w * 0.58)
        self.right_w = self.header_w - self.left_w - self.GAP
        self.BOTTOM_Y = self.TOP_Y + self.TOP_H + 20
        self.BOTTOM_H = self.H - self.BOTTOM_Y - self.M
        self.bottom_w = self.header_w
        self.cam_pad = 20
        self.cam_h = 540
        self.cam_w = self.left_w - (self.cam_pad * 2)

    # -----------------------------
    # Fonts (force Arial)
    # -----------------------------
    def setup_fonts(self):
        self.FONT_FAMILY = "Arial"
        for name in (
            "TkDefaultFont", "TkHeadingFont", "TkTextFont", "TkMenuFont",
            "TkFixedFont", "TkTooltipFont", "TkCaptionFont",
            "TkSmallCaptionFont", "TkIconFont"
        ):
            try:
                tkfont.nametofont(name).configure(family=self.FONT_FAMILY)
            except tk.TclError:
                pass

        self.F   = lambda size, bold=False: ctk.CTkFont(
            family=self.FONT_FAMILY, size=size, weight=("bold" if bold else "normal")
        )
        self.FTK = lambda size, bold=False: (
            (self.FONT_FAMILY, size, "bold") if bold else (self.FONT_FAMILY, size)
        )

    # -----------------------------
    # Data / Config
    # -----------------------------
    def initialize_data(self):
        # นับรูปทรงและ total
        self.shape_counts = {"heart": 0, "rectangle": 0, "circle": 0, "total": 0}

        # ตารางสถานะด้านล่าง (ค่าเริ่มต้น)
        self.defect_data = [
            ("รอยแตก", "ยังไม่พบ", "green"),
            ("รูเข็ม", "ยังไม่พบ", "green"),
        ]
        self.status_labels = {}
        self.defect_th_map = {"crack": "รอยแตก", "hole": "รูเข็ม"}
        self._defect_defaults = {d: (s, c) for d, s, c in self.defect_data}

        self.is_collecting_data = False
        self.camera_running = False
        self.cap = None

        # session & export
        self.session_rows = []
        self.session_meta = {}
        self.plate_id_counter = 1
        self.lot_id = self.generate_lot_id()
        self.lbl_lot = None

        # Path พื้นฐานและโฟลเดอร์บันทึก
        self.BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        self.save_root = os.path.join(self.BASE_DIR, "savefile")
        os.makedirs(self.save_root, exist_ok=True)

        # ---------- YOLO / Detection ----------
        # ตั้ง path โมเดลสอง-stage (ปรับชื่อไฟล์ตามของคุณ)
        self.SHAPE_WEIGHTS  = os.path.join(self.BASE_DIR, "models", "shape_best.pt")
        self.DEFECT_WEIGHTS = os.path.join(self.BASE_DIR, "models", "defect_best.pt")

        self.shape_model  = None
        self.defect_model = None

        # ค่าพื้นฐาน
        self.imgsz = 896

        # threshold แยกกันสำหรับสองโมเดล
        self.conf_shape = 0.30
        self.iou_shape  = 0.60
        self.conf_defect= 0.25
        self.iou_defect = 0.65

        # กลุ่มคลาส
        self.shape_classes_ultra = {
            "circle_leaf_plate",
            "heart_shaped_leaf_plate",
            "rectangular_leaf_plate"
        }
        self.defect_classes_ultra = {"crack", "hole"}

        # แมปชื่อ shape -> short
        self.shape_map = {
            "heart_shaped_leaf_plate": "heart",
            "rectangular_leaf_plate": "rectangle",
            "circle_leaf_plate": "circle",
        }
        # Display names for shapes in Thai
        self.shape_display_map = {
            "heart": "หัวใจ",
            "rectangle": "สี่เหลี่ยมผืนผ้า",
            "circle": "วงกลม",
        }

        # โฟลเดอร์รูป Annotated
        self.captures_dir = os.path.join(self.BASE_DIR, "captures")
        os.makedirs(self.captures_dir, exist_ok=True)
        self.save_cooldown_ms = 1200
        self._last_save_ms = 0

        # ไฟล์ CSV/JSON อัตโนมัติของ "รอบนี้"
        self._auto_csv_path = None
        self._auto_json_path = None
        self._session_stamp = None

        # ---------- Firebase ----------
        self.firebase_base = "https://leaf-plate-defect-detec-w-cnn-default-rtdb.asia-southeast1.firebasedatabase.app"
        self.firebase_session_key = None
        self._fb_ready = False  # Admin SDK พร้อมหรือยัง

        # UI labels (จะ set ใน create_xxx)
        self.lbl_heart = None
        self.lbl_rect  = None
        self.lbl_circle= None

        # Plate gating state
        self.gate_has_plate = False
        self.gate_has_counted = False
        self.gate_present_frames = 0
        self.gate_absent_frames = 0
        self.gate_present_thresh = 5
        self.gate_absent_thresh = 10

        # Plate status label placeholder
        self.lbl_plate_status = None

        # defect counts ต่อจาน (latched)
        self._latched_defect_counts = {"crack": 0, "hole": 0}

    # -----------------------------
    # Helpers for date/lot/defects
    # -----------------------------
    def generate_lot_id(self):
        return "PTP" + datetime.now().strftime("%y%m%d") + "_01"

    def thai_date(self, dt: datetime):
        return dt.strftime(f"%d/%m/{dt.year + 543}")

    def title_date(self, dt: datetime):
        return dt.strftime("%d/%m/%y")

    # -----------------------------
    # Firebase Admin helpers
    # -----------------------------
    def _fb_init(self):
        """Init Firebase Admin ด้วย service account 'serviceAccountKey.json' หนึ่งครั้ง"""
        if self._fb_ready:
            return
        try:
            cred_path = os.path.join(self.BASE_DIR, "serviceAccountKey.json")
            cred = credentials.Certificate(cred_path)
            firebase_admin.initialize_app(cred, {"databaseURL": self.firebase_base})
            self._fb_ready = True
        except Exception as e:
            print(f"[Firebase] Admin init failed, fallback to REST. reason={e}")
            self._fb_ready = False

    def _firebase_post(self, path, obj):
        """POST (push) : Admin SDK ก่อน, ไม่งั้น fallback REST"""
        payload = {
            **obj,
            "_meta": {
                "source": "python-admin",
                "pushed_at": datetime.now(timezone.utc).isoformat(),
                "server_id": str(uuid.uuid4())
            }
        }
        try:
            self._fb_init()
            if self._fb_ready:
                ref = db.reference(path)
                ref.push(payload)
                return
        except Exception as e:
            print(f"[Firebase] Admin push failed, fallback to REST. reason={e}")

        # --- Fallback: REST ---
        try:
            url = f"{self.firebase_base}/{path}.json"
            data = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=5):
                pass
        except Exception as e:
            print(f"[Firebase] REST POST error: {e}")

    def _firebase_put(self, path, obj):
        """PUT (set) : Admin SDK ก่อน, ไม่งั้น fallback REST"""
        try:
            self._fb_init()
            if self._fb_ready:
                ref = db.reference(path)
                ref.set(obj)
                return
        except Exception as e:
            print(f"[Firebase] Admin set failed, fallback to REST. reason={e}")

        # --- Fallback: REST ---
        try:
            url = f"{self.firebase_base}/{path}.json"
            data = json.dumps(obj).encode("utf-8")
            req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"}, method="PUT")
            with urllib.request.urlopen(req, timeout=5):
                pass
        except Exception as e:
            print(f"[Firebase] REST PUT error: {e}")

    # -----------------------------
    # App / Camera / Models
    # -----------------------------
    def setup_app(self):
        self.set_layout_constants()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.app = ctk.CTk()
        self.app.title("Leaf Plate Defect Detection - Two Stage (Full HD)")
        self.app.geometry(f"{self.W}x{self.H}+0+0")
        self.app.resizable(False, False)
        self.app.configure(fg_color="#ffffff")
        self.app.protocol("WM_DELETE_WINDOW", self.on_closing)

        # map Ctrl+C (SIGINT) ให้มาปิดหน้าต่างอย่างนุ่มนวล
        try:
            signal.signal(signal.SIGINT, lambda sig, frm: self.safe_after(0, self.on_closing))
        except Exception:
            pass

    def setup_camera(self):
        if sys.platform.startswith("win"):
            backend = cv2.CAP_DSHOW
        elif sys.platform == "darwin":
            backend = cv2.CAP_AVFOUNDATION
        else:
            backend = cv2.CAP_V4L2

        self.cap = cv2.VideoCapture(0, backend)
        if not self.cap or not self.cap.isOpened():
            self.cap = cv2.VideoCapture(0)
        if not self.cap or not self.cap.isOpened():
            print("Cannot open camera"); self.cap = None; return

        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

    def setup_models(self):
        # โหลดโมเดล shape
        try:
            assert os.path.exists(self.SHAPE_WEIGHTS), f"ไม่พบ shape weights: {self.SHAPE_WEIGHTS}"
            self.shape_model = YOLO(self.SHAPE_WEIGHTS)
        except Exception as e:
            messagebox.showerror("Model Error (Shape)", f"โหลดโมเดลรูปทรงไม่สำเร็จ:\n{e}")
            self.shape_model = None

        # โหลดโมเดล defect
        try:
            assert os.path.exists(self.DEFECT_WEIGHTS), f"ไม่พบ defect weights: {self.DEFECT_WEIGHTS}"
            self.defect_model = YOLO(self.DEFECT_WEIGHTS)
        except Exception as e:
            messagebox.showerror("Model Error (Defect)", f"โหลดโมเดลตำหนิไม่สำเร็จ:\n{e}")
            self.defect_model = None

        if (self.shape_model is None) or (self.defect_model is None):
            messagebox.showwarning("Warning", "ต้องโหลดโมเดลครบทั้ง shape และ defect ก่อนเริ่มทำงาน")

    # -----------------------------
    # UI
    # -----------------------------
    def create_widgets(self):
        self.create_header()
        self.create_main_content()

    def create_header(self):
        header_frame = ctk.CTkFrame(
            self.app, width=self.header_w, height=self.HEADER_H,
            fg_color="#7A5429", corner_radius=10
        )
        header_frame.place(x=self.M, y=self.HEADER_Y)

        ctk.CTkLabel(
            header_frame,
            text="โปรแกรมตรวจจับรูปทรง + รอยตำหนิ (Two-Stage)",
            font=self.F(32, True),
            text_color="white"
        ).place(x=50, y=28)

        self.header_time_label = ctk.CTkLabel(
            header_frame,
            text=datetime.now().strftime("%H:%M:%S"),
            font=self.F(22, True),
            text_color="white"
        )
        self.header_time_label.place(x=self.header_w - 220, y=15)

        current_date = datetime.now()
        thai_year = current_date.year + 543
        date_str = current_date.strftime(f"%d/%m/{thai_year}")
        self.header_date_label = ctk.CTkLabel(
            header_frame, text=date_str, font=self.F(18), text_color="white"
        )
        self.header_date_label.place(x=self.header_w - 220, y=50)

        self._update_header_time()

    def create_main_content(self):
        self.create_left_panel()
        self.create_right_panel()
        self.create_bottom_panel()

    def create_left_panel(self):
        left_frame = ctk.CTkFrame(
            self.app, width=self.left_w, height=self.TOP_H,
            fg_color="#ffffff", corner_radius=15,
            border_width=2, border_color="#7A5429"
        )
        left_frame.place(x=self.M, y=self.TOP_Y)

        self.camera_frame = ctk.CTkFrame(
            left_frame, width=self.cam_w, height=self.cam_h,
            fg_color="#E4DFDA", corner_radius=10,
            border_width=1, border_color="#7A5429"
        )
        self.camera_frame.place(x=self.cam_pad, y=self.cam_pad)

        self.camera_label = tk.Label(
            self.camera_frame, text="Initializing Camera...",
            font=self.FTK(16), fg="black", bg="#7A5429"
        )
        self.camera_label.place(x=0, y=0, width=self.cam_w, height=self.cam_h)

        self.create_control_buttons(left_frame)

    def create_control_buttons(self, parent):
        button_frame_w = self.left_w - 2 * self.cam_pad
        button_frame = ctk.CTkFrame(parent, width=button_frame_w, height=60, fg_color="transparent")
        button_frame.place(x=self.cam_pad, y=self.cam_pad + self.cam_h + 15)

        self.toggle_button = ctk.CTkButton(
            button_frame, width=150, height=45, text="เริ่ม",
            font=self.F(16, True), text_color="#FFFFFF",
            fg_color="#253BFA", hover_color="#0F0D69",
            command=self.toggle_data_collection
        )
        self.toggle_button.place(x=0, y=5)

        export_button = ctk.CTkButton(
            button_frame, width=150, height=45, text="Export",
            font=self.F(16, True), text_color="#FFFFFF",
            fg_color="#5BCADD", hover_color="#58AEBD",
            command=self.show_export_dialog
        )
        export_button.place(x=button_frame_w - 150, y=5)

    def create_right_panel(self):
        right_x = self.M + self.left_w + self.GAP
        right_frame = ctk.CTkFrame(
            self.app, width=self.right_w, height=self.TOP_H,
            fg_color="#E4DFDA", corner_radius=15,
            border_width=2, border_color="#7A5429"
        )
        right_frame.place(x=right_x, y=self.TOP_Y)
        self.create_unified_panel(right_frame)

    def create_unified_panel(self, parent):
        panel_w = self.right_w - 40
        panel_h = self.TOP_H - 40

        panel = ctk.CTkFrame(
            parent, width=panel_w, height=panel_h,
            fg_color="#E4DFDA", corner_radius=15, border_width=0
        )
        panel.place(x=20, y=20)

        ctk.CTkLabel(panel, text="บันทึกได้ทั้งหมด", font=self.F(26, True), text_color="#1a2a3a")\
            .place(x=panel_w // 2, y=40, anchor="center")

        self.total_number_label = ctk.CTkLabel(
            panel, text=str(self.shape_counts["total"]),
            font=self.F(70, True), text_color="#e74c3c"
        )
        self.total_number_label.place(x=panel_w // 2, y=110, anchor="center")

        ctk.CTkLabel(panel, text="จานแต่ละรูปทรง", font=self.F(20, True), text_color="#1a2a3a")\
            .place(x=panel_w // 2, y=180, anchor="center")

        # Card size and layout
        box_w, box_h = 220, 110
        box_y = 220
        gap = (panel_w - 3 * box_w) // 4

        # Heart
        heart_box = ctk.CTkFrame(panel, width=box_w, height=box_h,
                                 fg_color="#ffffff", corner_radius=18)
        heart_box.place(x=gap, y=box_y)
        ctk.CTkLabel(heart_box, text="หัวใจ", font=self.F(18, True),
                     text_color="#e74c3c").place(x=box_w // 2, y=32, anchor="center")
        self.lbl_heart = ctk.CTkLabel(heart_box, text=str(self.shape_counts["heart"]),
                                      font=self.F(40, True), text_color="#1a2a3a")
        self.lbl_heart.place(x=box_w // 2, y=70, anchor="center")

        # Rectangle
        rect_box = ctk.CTkFrame(panel, width=box_w, height=box_h,
                                fg_color="#ffffff", corner_radius=18)
        rect_box.place(x=gap * 2 + box_w, y=box_y)
        ctk.CTkLabel(rect_box, text="สี่เหลี่ยมผืนผ้า", font=self.F(18, True),
                     text_color="#199129").place(x=box_w // 2, y=32, anchor="center")
        self.lbl_rect = ctk.CTkLabel(rect_box, text=str(self.shape_counts["rectangle"]),
                                     font=self.F(40, True), text_color="#1a2a3a")
        self.lbl_rect.place(x=box_w // 2, y=70, anchor="center")

        # Circle
        circle_box = ctk.CTkFrame(panel, width=box_w, height=box_h,
                                  fg_color="#ffffff", corner_radius=18)
        circle_box.place(x=gap * 3 + box_w * 2, y=box_y)
        ctk.CTkLabel(circle_box, text="วงกลม", font=self.F(18, True),
                     text_color="#2AA7B8").place(x=box_w // 2, y=32, anchor="center")
        self.lbl_circle = ctk.CTkLabel(circle_box, text=str(self.shape_counts["circle"]),
                                       font=self.F(40, True), text_color="#1a2a3a")
        self.lbl_circle.place(x=box_w // 2, y=70, anchor="center")

        # Summary
        summary_card_y = box_y + box_h + 30
        summary_card_h = panel_h - summary_card_y - 20
        summary_card_w = panel_w - 2 * gap

        summary_card = ctk.CTkFrame(panel, width=summary_card_w, height=summary_card_h,
                                    fg_color="#ffffff", corner_radius=18)
        summary_card.place(x=gap, y=summary_card_y)

        ctk.CTkLabel(summary_card, text="สรุปผลการตรวจ",
                     font=self.F(20, True), text_color="#1a2a3a")\
            .place(x=summary_card_w // 2, y=32, anchor="center")

        label_x = 60
        value_x = summary_card_w // 2 + 20
        start_y = 70
        row_height = 38

        self.lbl_plate_order = ctk.CTkLabel(summary_card, text=str(self.shape_counts["total"]),
                                            font=self.F(18, True), text_color="#1a2a3a")
        self.lbl_plate_order.place(x=value_x, y=start_y, anchor="w")
        ctk.CTkLabel(summary_card, text="ลำดับจาน:", font=self.F(16), text_color="#1a2a3a")\
            .place(x=label_x, y=start_y, anchor="w")

        self.lbl_plate_status = ctk.CTkLabel(summary_card, text="ยังไม่ได้ตรวจ",
                                             font=self.F(18, True), text_color="#888888")
        self.lbl_plate_status.place(x=value_x, y=start_y + row_height, anchor="w")
        ctk.CTkLabel(summary_card, text="สถานะ:", font=self.F(16), text_color="#1a2a3a")\
            .place(x=label_x, y=start_y + row_height, anchor="w")

        self.lbl_defect_count = ctk.CTkLabel(summary_card, text="ยังไม่พบ",
                                             font=self.F(20, True), text_color="#888888")
        self.lbl_defect_count.place(x=value_x, y=start_y + row_height * 2, anchor="w")
        ctk.CTkLabel(summary_card, text="จำนวนข้อบกพร่อง:", font=self.F(16), text_color="#1a2a3a")\
            .place(x=label_x, y=start_y + row_height * 2, anchor="w")

        self.lbl_lot = ctk.CTkLabel(summary_card, text=f"{self.lot_id}",
                                    font=self.F(16, True), text_color="#1a2a3a")
        self.lbl_lot.place(x=value_x, y=start_y + row_height * 3, anchor="w")
        ctk.CTkLabel(summary_card, text="รหัสล็อต:", font=self.F(16), text_color="#1a2a3a")\
            .place(x=label_x, y=start_y + row_height * 3, anchor="w")

    def create_bottom_panel(self):
        bottom_frame = ctk.CTkFrame(
            self.app, width=self.bottom_w, height=self.BOTTOM_H,
            fg_color="#ffffff", corner_radius=15,
            border_width=2, border_color="#7A5429"
        )
        bottom_frame.place(x=self.M, y=self.BOTTOM_Y)
        self.create_defect_table(bottom_frame)

    def create_defect_table(self, parent):
        table_w = self.bottom_w - 40
        table_h = self.BOTTOM_H - 40
        table_frame = ctk.CTkFrame(parent, width=table_w, height=table_h,
                                   fg_color="#ffffff", corner_radius=10)
        table_frame.place(x=20, y=20)

        header_frame = ctk.CTkFrame(table_frame, width=table_w - 10, height=44,
                                    fg_color="#7A5429", corner_radius=8)
        header_frame.place(x=5, y=5)

        col1_x = int((table_w - 10) * 0.32)
        col2_x = int((table_w - 10) * 0.76)

        ctk.CTkLabel(header_frame, text="ประเภทตำหนิ", font=self.F(16, True),
                     text_color="white").place(x=col1_x, y=22, anchor="center")
        ctk.CTkLabel(header_frame, text="สถานะตำหนิ", font=self.F(16, True),
                     text_color="white").place(x=col2_x, y=22, anchor="center")

        row_start_y = 55
        row_h = 36
        for i, (defect, status, color) in enumerate(self.defect_data):
            row_y = row_start_y + i * row_h
            row_color = "#ffffff" if i % 2 == 0 else "#E4DFDA"
            row_frame = ctk.CTkFrame(table_frame, width=table_w - 10, height=row_h,
                                     fg_color=row_color, corner_radius=0)
            row_frame.place(x=5, y=row_y)

            ctk.CTkLabel(row_frame, text=defect, font=self.F(15),
                         text_color="#1a2a3a").place(x=col1_x, y=row_h // 2, anchor="center")

            status_color = "#199129" if color == "green" else "#e74c3c"
            lbl = ctk.CTkLabel(row_frame, text=status, font=self.F(15, True),
                               text_color=status_color)
            lbl.place(x=col2_x, y=row_h // 2, anchor="center")

            self.status_labels[defect] = lbl

    # ----------------- Export helpers (UTF-8 BOM + Excel Safe) -----------------
    def _excel_safe(self, s: str) -> str:
        if s is None:
            return ""
        s = str(s).replace("\n", "\r\n")
        if s and s[0] in ("=", "+", "-", "@"):
            s = "'" + s
        return s

    def _unique_filename(self, base_path, stem, extension):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"{stem}_{ts}{extension}"
        full = os.path.join(base_path, fname)
        i = 2
        while os.path.exists(full):
            fname = f"{stem}_{ts}_{i}{extension}"
            full = os.path.join(base_path, fname)
            i += 1
        return full

    def _write_csv(self, directory) -> str:
        csv_path = self._unique_filename(directory, "Report", ".csv")
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([self._excel_safe(f"รายงานการตรวจจานใบไม้ วันที่ {self.title_date(datetime.now())}")])
            w.writerow(["วันที่", "เวลา", "Plate ID", "Lot ID", "รูปทรงจาน", "ตำหนิที่พบ", "หมายเหตุ"])
            for r in self.session_rows:
                defects_for_csv = self._excel_safe(r["defects"])
                w.writerow([
                    r["date"], r["time"], r["plate_id"], r["lot_id"],
                    r.get("shape", "-"), defects_for_csv, self._excel_safe(r["note"])
                ])
        return csv_path

    def _write_json(self, directory) -> str:
        json_path = self._unique_filename(directory, "Report", ".json")
        payload = {
            "report_title": f"รายงานการตรวจจานใบไม้ วันที่ {self.title_date(datetime.now())}",
            "lot_id": self.lot_id,
            "session": {
                "start_time": self.session_meta.get("start_time"),
                "end_time": datetime.now().strftime("%H:%M:%S")
            },
            "records": self.session_rows
        }
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(payload, jf, ensure_ascii=False, indent=2)
        return json_path

    def export_session_csv_and_json(self, directory):
        if not self.session_rows:
            messagebox.showwarning("Warning", "ยังไม่มีข้อมูลสำหรับบันทึก")
            return
        csv_path = self._write_csv(directory)
        json_path = self._write_json(directory)
        messagebox.showinfo("Success", f"บันทึกไฟล์เรียบร้อย\n{os.path.basename(csv_path)}\n{os.path.basename(json_path)}")
        self._reset_all_and_next_lot()

    def show_export_dialog(self):
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("Export Options")
        dialog.geometry("320x220")
        dialog.resizable(False, False)
        dialog.transient(self.app); dialog.grab_set()

        ctk.CTkLabel(dialog, text="Choose Export Format", font=self.F(16, True))\
            .place(x=160, y=30, anchor="center")

        ctk.CTkButton(
            dialog, width=220, height=38, text="Export to CSV", font=self.F(12, True),
            text_color="#FFFFFF", fg_color="#50C878", hover_color="#27ad60",
            command=lambda: self.save_to_csv(dialog)
        ).place(x=160, y=80, anchor="center")

        ctk.CTkButton(
            dialog, width=220, height=38, text="Export to JSON", font=self.F(12, True),
            text_color="#FFFFFF", fg_color="#f1c40f", hover_color="#d4ac0d",
            command=lambda: self.save_to_json(dialog)
        ).place(x=160, y=125, anchor="center")

        ctk.CTkButton(
            dialog, width=220, height=38, text="Export Both", font=self.F(12, True),
            text_color="#FFFFFF", fg_color="#7A5429", hover_color="#2980b9",
            command=lambda: self.save_collected_to_csv_and_json(dialog)
        ).place(x=160, y=170, anchor="center")

    def save_to_csv(self, dialog):
        if not self.session_rows:
            messagebox.showwarning("Warning", "ยังไม่มีข้อมูลสำหรับบันทึก")
            return
        directory = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึก CSV")
        if directory:
            csv_path = self._write_csv(directory)
            messagebox.showinfo("Success", f"บันทึกไฟล์เรียบร้อย\n{os.path.basename(csv_path)}")
            dialog.destroy()
            self._reset_all_and_next_lot()

    def save_to_json(self, dialog):
        if not self.session_rows:
            messagebox.showwarning("Warning", "ยังไม่มีข้อมูลสำหรับบันทึก")
            return
        directory = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึก JSON")
        if directory:
            json_path = self._write_json(directory)
            messagebox.showinfo("Success", f"บันทึกไฟล์เรียบร้อย\n{os.path.basename(json_path)}")
            dialog.destroy()
            self._reset_all_and_next_lot()

    def save_collected_to_csv_and_json(self, dialog):
        if not self.session_rows:
            messagebox.showwarning("Warning", "ยังไม่มีข้อมูลสำหรับบันทึก")
            return
        directory = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึกไฟล์รายงาน")
        if directory:
            csv_path = self._write_csv(directory)
            json_path = self._write_json(directory)
            messagebox.showinfo(
                "Success",
                f"บันทึกไฟล์เรียบร้อย\n{os.path.basename(csv_path)}\n{os.path.basename(json_path)}"
            )
            dialog.destroy()
            self._reset_all_and_next_lot()

    # ----------------- Auto-save (CSV/JSON in ./savefile) + Firebase -----------------
    def _ensure_session_files(self):
        if self._auto_csv_path and self._auto_json_path:
            return
        self._session_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self._auto_csv_path  = os.path.join(self.save_root, f"Report_{self._session_stamp}.csv")
        self._auto_json_path = os.path.join(self.save_root, f"Report_{self._session_stamp}.json")

        with open(self._auto_csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([self._excel_safe(f"รายงานการตรวจจานใบไม้ วันที่ {self.title_date(datetime.now())}")])
            w.writerow(["วันที่", "เวลา", "Plate ID", "Lot ID", "รูปทรงจาน", "ตำหนิที่พบ", "หมายเหตุ"])

        self.session_meta.setdefault("start_time", datetime.now().strftime("%H:%M:%S"))

        self.firebase_session_key = self._session_stamp  # ใช้ stamp เป็นชื่อ session
        self._write_json_to_path(self._auto_json_path)

    def _append_csv_row_to_path(self, row):
        with open(self._auto_csv_path, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            defects_for_csv = self._excel_safe(row["defects"])
            w.writerow([
                row["date"], row["time"], row["plate_id"], row["lot_id"],
                row.get("shape", "-"), defects_for_csv, self._excel_safe(row["note"])
            ])

    def _write_json_to_path(self, path):
        payload = {
            "report_title": f"รายงานการตรวจจานใบไม้ วันที่ {self.title_date(datetime.now())}",
            "lot_id": self.lot_id,
            "session": {
                "start_time": self.session_meta.get("start_time"),
                "end_time": datetime.now().strftime("%H:%M:%S")
            },
            "records": self.session_rows
        }
        try:
            with open(path, "w", encoding="utf-8") as jf:
                json.dump(payload, jf, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Write JSON error: {e}")

    def _append_csv_json_and_firebase(self, row):
        self._ensure_session_files()
        self._append_csv_row_to_path(row)
        self._write_json_to_path(self._auto_json_path)

        meta = {
            "report_title": f"รายงานการตรวจจานใบไม้ วันที่ {self.title_date(datetime.now())}",
            "lot_id": self.lot_id,
            "session": {
                "start_time": self.session_meta.get("start_time"),
                "end_time": datetime.now().strftime("%H:%M:%S")
            }
        }
        self._firebase_put(f"sessions/{self.firebase_session_key}/meta", meta)
        self._firebase_post(f"sessions/{self.firebase_session_key}/records", row)

    # ----------------- Lot/Reset helpers -----------------
    def _update_lot_label(self):
        try:
            if self.lbl_lot is not None:
                self.lbl_lot.configure(text=f"รหัสล็อต : {self.lot_id}")
        except Exception:
            pass

    def _increment_lot_id(self):
        try:
            now_date_short = datetime.now().strftime("%y%m%d")
            base = f"PTP{now_date_short}"
            seq = 1
            if isinstance(self.lot_id, str) and "_" in self.lot_id and self.lot_id.startswith("PTP"):
                old_base, old_seq = self.lot_id.split("_", 1)
                if old_base.endswith(now_date_short):
                    try:
                        seq = int(old_seq) + 1
                    except Exception:
                        seq = 1
            self.lot_id = f"{base}_{seq:02d}"
        except Exception:
            self.lot_id = self.generate_lot_id()
        self._update_lot_label()

    def _reset_all_and_next_lot(self):
        self.shape_counts = {"heart": 0, "rectangle": 0, "circle": 0, "total": 0}
        try:
            if self.lbl_heart:  self.lbl_heart.configure(text="0")
            if self.lbl_rect:   self.lbl_rect.configure(text="0")
            if self.lbl_circle: self.lbl_circle.configure(text="0")
            self.total_number_label.configure(text="0")
            if self.lbl_plate_order: self.lbl_plate_order.configure(text="0")
            if self.lbl_defect_count: self.lbl_defect_count.configure(text="ยังไม่พบ", text_color="#888888")
            if self.lbl_plate_status: self.lbl_plate_status.configure(text="ยังไม่ได้ตรวจ", text_color="#888888")
        except Exception:
            pass

        self.session_rows = []
        self.session_meta = {}
        self.plate_id_counter = 1
        if hasattr(self, "lbl_plate_no") and self.lbl_plate_no is not None:
            try:
                self.lbl_plate_no.configure(text="จานที่ : -")
            except Exception:
                pass

        self._auto_csv_path = None
        self._auto_json_path = None
        self._session_stamp = None
        self.firebase_session_key = None

        self.gate_has_plate = False
        self.gate_has_counted = False
        self.gate_present_frames = 0
        self.gate_absent_frames = 0
        self._set_plate_status("pending")
        try:
            self._reset_defect_table()
            self._render_latched_defect_counts()
        except Exception:
            pass

        self._increment_lot_id()

    # ----------------- Detection helpers -----------------
    def _update_shape_counters(self, shapes_found: set):
        if not shapes_found:
            return
        for shp in shapes_found:
            if shp in self.shape_counts:
                self.shape_counts[shp] += 1
                self.shape_counts["total"] += 1

        if self.lbl_plate_order:
            self.lbl_plate_order.configure(text=str(self.shape_counts["total"]))

        if self.lbl_heart:  self.lbl_heart.configure(text=str(self.shape_counts["heart"]))
        if self.lbl_rect:   self.lbl_rect.configure(text=str(self.shape_counts["rectangle"]))
        if self.lbl_circle: self.lbl_circle.configure(text=str(self.shape_counts["circle"]))
        self.total_number_label.configure(text=str(self.shape_counts["total"]))

    def _update_defect_status_ui(self, defect_names: set):
        for en_name in defect_names:
            th_name = self.defect_th_map.get(en_name)
            if not th_name:
                continue
            lbl = self.status_labels.get(th_name)
            if lbl:
                lbl.configure(text="พบตำหนิ", text_color="#e74c3c")

    def _update_defect_counts_ui(self, defect_counts: dict):
        for en_name, th_name in self.defect_th_map.items():
            lbl = self.status_labels.get(th_name)
            if not lbl:
                continue
            cnt = int(defect_counts.get(en_name, 0))
            if cnt > 0:
                lbl.configure(text=str(cnt), text_color="#e74c3c")
            else:
                default_status, default_color = self._defect_defaults.get(th_name, ("ยังไม่พบ", "green"))
                status_color = "#199129" if default_color == "green" else "#e74c3c"
                lbl.configure(text=default_status, text_color=status_color)

    def _reset_defect_table(self):
        for k in list(self._latched_defect_counts.keys()):
            self._latched_defect_counts[k] = 0
        for defect, (status, color) in self._defect_defaults.items():
            lbl = self.status_labels.get(defect)
            if lbl:
                status_color = "#199129" if color == "green" else "#e74c3c"
                lbl.configure(text=status, text_color=status_color)

    def _render_latched_defect_counts(self):
        for en_name, th_name in self.defect_th_map.items():
            lbl = self.status_labels.get(th_name)
            if not lbl:
                continue
            cnt = int(self._latched_defect_counts.get(en_name, 0))
            if cnt > 0:
                lbl.configure(text=str(cnt), text_color="#e74c3c")
            else:
                default_status, default_color = self._defect_defaults.get(th_name, ("ยังไม่พบ", "green"))
                status_color = "#199129" if default_color == "green" else "#e74c3c"
                lbl.configure(text=default_status, text_color=status_color)

    def _set_plate_status(self, mode: str, defect_count=None):
        if not self.lbl_plate_status:
            return
        if mode == "pending":
            self.lbl_plate_status.configure(text="ยังไม่ได้ตรวจ", text_color="#888888")
            self.lbl_defect_count.configure(text="ยังไม่พบ", text_color="#888888")
        elif mode == "pass":
            self.lbl_plate_status.configure(text="ผ่าน", text_color="#199129")
            self.lbl_defect_count.configure(text="ไม่มีตำหนิ", text_color="#888888")
        elif mode == "fail":
            self.lbl_plate_status.configure(text="มีตำหนิ", text_color="#e74c3c")
            if defect_count is not None and defect_count > 0:
                self.lbl_defect_count.configure(text=str(defect_count), text_color="#e74c3c")
            else:
                self.lbl_defect_count.configure(text="ไม่มีตำหนิ", text_color="#888888")
        elif mode == "counted":
            if defect_count is not None and defect_count > 0:
                self.lbl_plate_status.configure(text="มีตำหนิ", text_color="#e74c3c")
                self.lbl_defect_count.configure(text=str(defect_count), text_color="#e74c3c")
            else:
                self.lbl_plate_status.configure(text="ผ่าน", text_color="#199129")
                self.lbl_defect_count.configure(text="ไม่มีตำหนิ", text_color="#888888")

    def _save_detection_record(self, annotated_bgr, defect_names: set, shapes_found: set):
        now = datetime.now()
        ts  = now.strftime("%Y%m%d_%H%M%S_%f")[:-3]

        img_path = os.path.join(self.captures_dir, f"detect_{ts}.jpg")
        try:
            cv2.imwrite(img_path, annotated_bgr)
        except Exception as e:
            print(f"Save image error: {e}")

        defects_th = [self.defect_th_map.get(d, d) for d in sorted(defect_names)]
        defects_text = " - " if not defects_th else " / ".join(defects_th)

        if not shapes_found:
            shape_text = "-"
        else:
            thai_shapes = [self.shape_display_map.get(s, s) for s in sorted(shapes_found)]
            shape_text = " / ".join(thai_shapes)

        row = {
            "date": self.thai_date(now),
            "time": now.strftime("%H:%M:%S"),
            "plate_id": self.plate_id_counter,
            "lot_id": self.lot_id,
            "shape": shape_text,
            "defects": defects_text.strip(),
            "note": ""
        }
        self.plate_id_counter += 1
        self.session_rows.append(row)

        self._update_defect_status_ui(defect_names)
        return row

    def _annotate_and_summarize_two_stage(self, frame_bgr, shape_res, defect_res):
        """
        รวมผลสองโมเดลในเฟรมเดียว → annotate + สรุป set ต่างๆ
        """
        annotated = frame_bgr.copy()
        shapes_found, defect_names = set(), set()
        defect_counts = {}

        # stage-1: shapes
        if shape_res is not None and hasattr(shape_res, "boxes") and shape_res.boxes is not None:
            names = self.shape_model.names if self.shape_model else {}
            boxes = shape_res.boxes
            xyxy = boxes.xyxy.cpu().numpy().astype(int) if boxes.xyxy is not None else []
            clss = boxes.cls.cpu().numpy().astype(int)   if boxes.cls is not None else []
            conf = boxes.conf.cpu().numpy()              if boxes.conf is not None else []

            for (x1, y1, x2, y2), c, p in zip(xyxy, clss, conf):
                label = names.get(int(c), str(c))
                # เฉพาะคลาสรูปทรงเท่านั้น
                if label in self.shape_classes_ultra:
                    cv2.rectangle(annotated, (x1, y1), (x2, y2), (0, 140, 255), 2)
                    cv2.putText(annotated, f"{label} {p:.2f}", (x1, max(20, y1 - 6)),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (10, 90, 255), 2, cv2.LINE_AA)
                    short = self.shape_map.get(label, label)
                    shapes_found.add(short)

        # stage-2: defects
        if defect_res is not None and hasattr(defect_res, "boxes") and defect_res.boxes is not None:
            names = self.defect_model.names if self.defect_model else {}
            boxes = defect_res.boxes
            xyxy = boxes.xyxy.cpu().numpy().astype(int) if boxes.xyxy is not None else []
            clss = boxes.cls.cpu().numpy().astype(int)   if boxes.cls is not None else []
            conf = boxes.conf.cpu().numpy()              if boxes.conf is not None else []

            for (x1, y1, x2, y2), c, p in zip(xyxy, clss, conf):
                label = names.get(int(c), str(c))
                if label in self.defect_classes_ultra:
                    cv2.rectangle(annotated, (x1, y1), (x2, y2), (255, 40, 40), 2)
                    cv2.putText(annotated, f"{label} {p:.2f}", (x1, max(20, y1 - 6)),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (20, 20, 255), 2, cv2.LINE_AA)
                    defect_names.add(label)
                    defect_counts[label] = defect_counts.get(label, 0) + 1

        return annotated, shapes_found, defect_counts, defect_names

    # ----------------- Camera loop (safe_after) -----------------
    def safe_after(self, delay_ms, func):
        """เรียก Tk.after แบบปลอดภัย: ไม่ยิงต่อถ้า window ถูกปิด"""
        try:
            if self.app is not None and self.app.winfo_exists():
                self.app.after(delay_ms, func)
        except tk.TclError:
            pass

    def start_camera(self):
        if self.cap:
            self.camera_running = True
            self._update_camera()

    def stop_camera(self):
        self.camera_running = False
        if self.cap:
            try:
                self.cap.release()
            except Exception:
                pass
            self.cap = None

    def _update_camera(self):
        if not self.camera_running or not self.cap:
            return

        ret, frame = self.cap.read()
        if not ret:
            self.safe_after(30, self._update_camera)
            return

        frame_resized = cv2.resize(frame, (self.cam_w, self.cam_h))
        frame_to_show = frame_resized

        if self.is_collecting_data and (self.shape_model is not None) and (self.defect_model is not None):
            try:
                # Stage-1 (shape)
                shape_results = self.shape_model.predict(
                    source=frame_resized,
                    imgsz=self.imgsz,
                    conf=self.conf_shape,
                    iou=self.iou_shape,
                    verbose=False
                )
                shape_res = shape_results[0]

                # Stage-2 (defect)
                defect_results = self.defect_model.predict(
                    source=frame_resized,
                    imgsz=self.imgsz,
                    conf=self.conf_defect,
                    iou=self.iou_defect,
                    verbose=False
                )
                defect_res = defect_results[0]

                annotated, shapes_found, defect_counts, defect_names = \
                    self._annotate_and_summarize_two_stage(frame_resized, shape_res, defect_res)

                frame_to_show = annotated

                # อัปเดตตาราง defect ด้วยจำนวน defect ล่าสุด
                self._update_defect_counts_ui(defect_counts)

                # Gating per plate
                plate_detected = (len(shapes_found) > 0) or (len(defect_names) > 0)

                # latched defect counts ต่อจาน
                if plate_detected:
                    for k, v in defect_counts.items():
                        try:
                            iv = int(v)
                        except Exception:
                            iv = 0
                        self._latched_defect_counts[k] = max(self._latched_defect_counts.get(k, 0), iv)
                    self._render_latched_defect_counts()

                if plate_detected:
                    self.gate_present_frames += 1
                    self.gate_absent_frames = 0
                else:
                    self.gate_absent_frames += 1
                    self.gate_present_frames = 0

                # New stable plate appears
                if (not self.gate_has_plate) and plate_detected and self.gate_present_frames >= self.gate_present_thresh:
                    self.gate_has_plate = True
                    self.gate_has_counted = False
                    self._set_plate_status("pending")
                    self._reset_defect_table()
                    self._render_latched_defect_counts()

                # Count once per plate
                if self.gate_has_plate and (not self.gate_has_counted) and self.gate_present_frames >= self.gate_present_thresh:
                    self._update_shape_counters(shapes_found)
                    if (not shapes_found) and (len(defect_names) > 0):
                        self.shape_counts["total"] += 1
                        self.total_number_label.configure(text=str(self.shape_counts["total"]))

                    row = self._save_detection_record(annotated, defect_names, shapes_found)
                    self._append_csv_json_and_firebase(row)

                    defect_count = sum(defect_counts.values())
                    self._set_plate_status("counted", defect_count)
                    self.gate_has_counted = True
                    self._last_save_ms = time.time() * 1000.0

                    if hasattr(self, "lbl_plate_no") and self.lbl_plate_no is not None:
                        try:
                            self.lbl_plate_no.configure(text=f"จานที่ : {row['plate_id']}")
                        except Exception:
                            pass

                # Plate removed -> reset gate
                if self.gate_has_plate and self.gate_absent_frames >= self.gate_absent_thresh:
                    self.gate_has_plate = False
                    self.gate_has_counted = False
                    self._set_plate_status("pending")

            except Exception as e:
                print(f"Inference error: {e}")
                frame_to_show = frame_resized

        # แสดงภาพ
        try:
            frame_rgb = cv2.cvtColor(frame_to_show, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(frame_rgb)
            imgtk = ImageTk.PhotoImage(pil_img)
            self.camera_label.configure(image=imgtk, text="")
            self.camera_label.image = imgtk
        except Exception as e:
            print(f"Camera display error: {e}")

        self.safe_after(30, self._update_camera)

    # ----------------- Stop confirm popup -----------------
    def stop_and_finalize(self):
        """หยุดการตรวจ + export อัตโนมัติลง ./savefile/ + อัปเดต meta ไป Firebase"""
        self.is_collecting_data = False
        try:
            self.toggle_button.configure(text="เริ่ม", fg_color="#297A3A")
        except Exception:
            pass

        if self.session_rows:
            try:
                self._write_csv(self.save_root)
                self._write_json(self.save_root)
                self._reset_all_and_next_lot()
                try:
                    messagebox.showinfo("Reset", f"รีเซ็ตข้อมูลเรียบร้อย\nรหัสล็อตใหม่: {self.lot_id}")
                except Exception:
                    pass
            except Exception as e:
                print(f"[Export at stop] failed: {e}")

        if self._session_stamp:
            meta = {
                "report_title": f"รายงานการตรวจจานใบไม้ วันที่ {self.title_date(datetime.now())}",
                "lot_id": self.lot_id,
                "session": {
                    "start_time": self.session_meta.get("start_time"),
                    "end_time": datetime.now().strftime("%H:%M:%S")
                }
            }
            self._firebase_put(f"sessions/{self.firebase_session_key}/meta", meta)
        self._set_plate_status("pending")

    def show_stop_confirm_dialog(self):
        dlg = ctk.CTkToplevel(self.app)
        dlg.title("ยืนยันการหยุด")
        dlg.geometry("380x190")
        dlg.resizable(False, False)
        dlg.transient(self.app)
        dlg.grab_set()

        try:
            self.app.update_idletasks()
            ax, ay = self.app.winfo_x(), self.app.winfo_y()
            aw, ah = self.app.winfo_width(), self.app.winfo_height()
            dw, dh = 380, 190
            dlg.geometry(f"{dw}x{dh}+{ax + (aw-dw)//2}+{ay + (ah-dh)//2}")
        except Exception:
            pass

        ctk.CTkLabel(dlg, text="คุณยืนยันจะหยุดตรวจและบันทึกรายงานหรือไม่?", font=self.F(16, True))\
            .place(x=190, y=55, anchor="center")

        ctk.CTkButton(
            dlg, width=140, height=40, text="Submit (ยืนยัน)",
            font=self.F(14, True), fg_color="#27ae60", hover_color="#1e8449",
            command=lambda: self._handle_stop_submit(dlg)
        ).place(x=40, y=110)

        ctk.CTkButton(
            dlg, width=140, height=40, text="ตรวจต่อ",
            font=self.F(14, True), fg_color="#95a5a6", hover_color="#7f8c8d",
            command=dlg.destroy
        ).place(x=200, y=110)

        dlg.bind("<Return>", lambda e: self._handle_stop_submit(dlg))
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def _handle_stop_submit(self, dlg):
        try:
            dlg.destroy()
        except Exception:
            pass
        self.stop_and_finalize()

    # ----------------- Events -----------------
    def toggle_data_collection(self):
        if not self.is_collecting_data:
            if (self.shape_model is None) or (self.defect_model is None):
                messagebox.showerror("Model Error", "ยังโหลดโมเดลไม่ครบ (shape/defect)")
                return
            self.is_collecting_data = True
            self.session_meta.setdefault("start_time", datetime.now().strftime("%H:%M:%S"))
            self.toggle_button.configure(text="หยุด", fg_color="#e74c3c", hover_color="#c0392b")

            # Reset gating state when starting
            self.gate_has_plate = False
            self.gate_has_counted = False
            self.gate_present_frames = 0
            self.gate_absent_frames = 0
            self._set_plate_status("pending")

            try:
                self._reset_defect_table()
                if hasattr(self, "lbl_plate_no") and self.lbl_plate_no is not None:
                    self.lbl_plate_no.configure(text=f"จานที่ : {self.plate_id_counter - 1}")
            except Exception:
                pass
        else:
            self.show_stop_confirm_dialog()

    def _update_header_time(self):
        try:
            self.header_time_label.configure(text=datetime.now().strftime("%H:%M:%S"))
        except Exception:
            pass
        self.safe_after(1000, self._update_header_time)

    def on_closing(self):
        if getattr(self, "is_collecting_data", False):
            try:
                if not messagebox.askyesno("ออกจากโปรแกรม", "กำลังตรวจจับอยู่ ต้องการออกทันทีหรือไม่?"):
                    return
            except Exception:
                pass
            self.stop_and_finalize()

        self.stop_camera()
        try:
            self.app.destroy()
        except Exception:
            pass

    # ----------------- Run -----------------
    def run(self):
        self.app.mainloop()


# ----------------- Boot -----------------
if __name__ == "__main__":
    app = LeafPlateTwoStageApp()
    app.initialize_data()
    app.setup_app()
    app.setup_fonts()
    app.setup_camera()
    app.setup_models()
    app.create_widgets()
    app.start_camera()
    try:
        app.run()
    except KeyboardInterrupt:
        # ปิดให้เรียบร้อย ถ้ามี Ctrl+C จากคอนโซล
        try:
            app.on_closing()
        except Exception:
            pass
