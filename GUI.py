import customtkinter as ctk
import tkinter as tk
from tkinter import font as tkfont
import cv2
from PIL import Image, ImageTk
from datetime import datetime, timedelta
import os, sys, json, csv, random
from tkinter import filedialog, messagebox


class LeafPlateDetectionApp:
    """
    Leaf Plate Defect Detection Application
    Full HD (1920x1080) layout; bigger camera; shorter Current Session card;
    table with larger fonts but compact height. Default font = Arial (forced).
    """

    def __init__(self):
        self.initialize_data()
        self.setup_app()
        self.setup_fonts()
        self.setup_camera()
        self.create_widgets()
        self.create_mock_data()
        self.start_camera()

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
    # Data
    # -----------------------------
    def initialize_data(self):
        self.shape_counts = {"heart": 28, "rectangle": 31, "circle": 10, "total": 81}
        self.defect_data = [
            ("รอยแตก", "ผ่าน", "green"),
            ("รอยยุบหรือรอยพับ", "พบตำหนิ", "red"),
            ("จุดใหญ่หรือสีคล้ำ", "ผ่าน", "green"),
            ("รูพรุนหรือรูเข็บ", "ผ่าน", "green"),
            ("รอยขีดข่วนหรือรอยกลอก", "ผ่าน", "green")
        ]
        self.is_collecting_data = False
        self.camera_running = False
        self.cap = None

        # session & export
        self.session_rows = []
        self.session_meta = {}
        self.collect_job_id = None
        self.collect_interval_ms = 1000  # ความถี่การเพิ่มแถว (ms) — ปรับได้
        self.plate_id_counter = 1
        self.last_log_time = None  # ไม่ใช้เร่งเวลาแล้ว แต่เก็บล่าสุดเผื่อโชว์
        self.lot_id = self.generate_lot_id()

        # mock เดิม (คงไว้)
        self.mock_json_data = {
            "leaf_plate_reports": {
                "session_1": {
                    "plates": {
                        "plate_001": {"time": "14:30:15", "defects": ["รอยแตก"], "shape": "heart"},
                        "plate_002": {"time": "14:31:20", "defects": [], "shape": "rectangle"},
                        "plate_003": {"time": "14:32:25", "defects": ["รอยยุบ"], "shape": "circle"}
                    }
                }
            }
        }

    # -----------------------------
    # Helpers for date/lot/defects
    # -----------------------------
    def generate_lot_id(self):
        return "PTP" + datetime.now().strftime("%y%m%d") + "_01"

    def thai_date(self, dt: datetime):
        return dt.strftime(f"%d/%m/{dt.year + 543}")

    def title_date(self, dt: datetime):
        return dt.strftime("%d/%m/%y")

    def random_defects_for_cell(self):
        """ใช้ bullet '• ' เพื่อไม่ให้ Excel คิดว่าเป็นสูตร (#NAME?)"""
        candidates = [
            "รอยย่นหรือรอยพับ",
            "จุดใหญ่หรือสีคล้ำ",
            "จุดไหม้",
            "รอยแตก",
            "รูพรุนหรือรูเข็บ"
        ]
        if random.random() < 0.5:
            return "-"   # ไม่มีตำหนิ
        k = 1 if random.random() < 0.7 else 2
        picked = random.sample(candidates, k=k)
        return "\n".join([f"• {p}" for p in picked])

    # ป้องกัน Excel ตีความเป็นสูตร + แทนบรรทัดใหม่เป็น \r\n
    def _excel_safe(self, s: str) -> str:
        if s is None:
            return ""
        s = str(s).replace("\n", "\r\n")
        if s and s[0] in ("=", "+", "-", "@"):
            s = "'" + s
        return s

    def next_mock_row(self):
        """
        ใช้เวลาจริงจากเครื่อง ณ ตอนสร้างแถว (ไม่เร่งเวลา)
        """
        now = datetime.now()
        self.last_log_time = now  # เก็บไว้เป็นข้อมูลล่าสุด
        row = {
            "date": self.thai_date(now),
            "time": now.strftime("%H:%M:%S"),
            "plate_id": self.plate_id_counter,
            "lot_id": self.lot_id,
            "defects": self.random_defects_for_cell(),
            "note": ""
        }
        self.plate_id_counter += 1
        return row

    # -----------------------------
    # App / Camera
    # -----------------------------
    def setup_app(self):
        self.set_layout_constants()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.app = ctk.CTk()
        self.app.title("Leaf Plate Defect Detection - Full HD")
        self.app.geometry(f"{self.W}x{self.H}+0+0")
        self.app.resizable(False, False)
        self.app.configure(fg_color="#ffffff")
        self.app.protocol("WM_DELETE_WINDOW", self.on_closing)

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

    # -----------------------------
    # UI
    # -----------------------------
    def create_widgets(self):
        self.create_header()
        self.create_main_content()

    def create_header(self):
        header_frame = ctk.CTkFrame(
            self.app, width=self.header_w, height=self.HEADER_H,
            fg_color="#1a2a3a", corner_radius=10
        )
        header_frame.place(x=self.M, y=self.HEADER_Y)

        ctk.CTkLabel(
            header_frame,
            text="Leaf Plate Defect Detection System",
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

        self.update_header_time()

    def create_main_content(self):
        self.create_left_panel()
        self.create_right_panel()
        self.create_bottom_panel()

    def create_left_panel(self):
        left_frame = ctk.CTkFrame(
            self.app, width=self.left_w, height=self.TOP_H,
            fg_color="#ffffff", corner_radius=15,
            border_width=2, border_color="#aed6f1"
        )
        left_frame.place(x=self.M, y=self.TOP_Y)

        self.camera_frame = ctk.CTkFrame(
            left_frame, width=self.cam_w, height=self.cam_h,
            fg_color="#e0f2f7", corner_radius=10,
            border_width=1, border_color="#aed6f1"
        )
        self.camera_frame.place(x=self.cam_pad, y=self.cam_pad)

        self.camera_label = tk.Label(
            self.camera_frame, text="Initializing Camera...",
            font=self.FTK(16), fg="black", bg="#e0f2f7"
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
            fg_color="#3498db", hover_color="#FFD700",
            command=self.toggle_data_collection
        )
        self.toggle_button.place(x=0, y=5)

        export_button = ctk.CTkButton(
            button_frame, width=150, height=45, text="Export",
            font=self.F(16, True), text_color="#FFFFFF",
            fg_color="#50C878", hover_color="#27ad60",
            command=self.show_export_dialog
        )
        export_button.place(x=button_frame_w - 150, y=5)

    def create_right_panel(self):
        right_x = self.M + self.left_w + self.GAP
        right_frame = ctk.CTkFrame(
            self.app, width=self.right_w, height=self.TOP_H,
            fg_color="#ffffff", corner_radius=15,
            border_width=2, border_color="#aed6f1"
        )
        right_frame.place(x=right_x, y=self.TOP_Y)

        self.create_total_count_card(right_frame)
        self.create_shape_counts_card(right_frame)
        self.create_session_info_card(right_frame)

    def create_total_count_card(self, parent):
        card_w = self.right_w - 40
        total_card = ctk.CTkFrame(
            parent, width=card_w, height=150,
            fg_color="#e0f2f7", corner_radius=10,
            border_width=2, border_color="#3498db"
        )
        total_card.place(x=20, y=20)

        ctk.CTkLabel(
            total_card, text="นับได้ทั้งหมด",
            font=self.F(18, True), text_color="#1a2a3a"
        ).place(x=card_w // 2, y=24, anchor="center")

        self.total_number_label = ctk.CTkLabel(
            total_card, text=str(self.shape_counts["total"]),
            font=self.F(56, True), text_color="#e74c3c"
        )
        self.total_number_label.place(x=card_w // 2, y=88, anchor="center")

    def create_shape_counts_card(self, parent):
        card_w = self.right_w - 40
        shape_card = ctk.CTkFrame(
            parent, width=card_w, height=200,
            fg_color="#e0f2f7", corner_radius=10,
            border_width=2, border_color="#3498db"
        )
        shape_card.place(x=20, y=190)

        ctk.CTkLabel(
            shape_card, text="จานแต่ละรูปแบบ",
            font=self.F(18, True), text_color="#1a2a3a"
        ).place(x=card_w // 2, y=20, anchor="center")

        box_w, box_h = 220, 100
        gap = max(12, (card_w - 3 * box_w) // 4)
        y0 = 55

        heart_frame = ctk.CTkFrame(shape_card, width=box_w, height=box_h, fg_color="#ffffff", corner_radius=8)
        heart_frame.place(x=gap, y=y0)
        ctk.CTkLabel(heart_frame, text="Heart", font=self.F(14, True), text_color="#e74c3c")\
            .place(x=box_w // 2, y=30, anchor="center")
        ctk.CTkLabel(heart_frame, text=str(self.shape_counts["heart"]), font=self.F(22, True),
                     text_color="#1a2a3a").place(x=box_w // 2, y=70, anchor="center")

        rect_frame = ctk.CTkFrame(shape_card, width=box_w, height=box_h, fg_color="#ffffff", corner_radius=8)
        rect_frame.place(x=gap * 2 + box_w, y=y0)
        ctk.CTkLabel(rect_frame, text="Rectangle", font=self.F(14, True), text_color="#3498db")\
            .place(x=box_w // 2, y=30, anchor="center")
        ctk.CTkLabel(rect_frame, text=str(self.shape_counts["rectangle"]), font=self.F(22, True),
                     text_color="#1a2a3a").place(x=box_w // 2, y=70, anchor="center")

        circle_frame = ctk.CTkFrame(shape_card, width=box_w, height=box_h, fg_color="#ffffff", corner_radius=8)
        circle_frame.place(x=gap * 3 + box_w * 2, y=y0)
        ctk.CTkLabel(circle_frame, text="Circle", font=self.F(14, True), text_color="#199129")\
            .place(x=box_w // 2, y=30, anchor="center")
        ctk.CTkLabel(circle_frame, text=str(self.shape_counts["circle"]), font=self.F(22, True),
                     text_color="#1a2a3a").place(x=box_w // 2, y=70, anchor="center")

    def create_session_info_card(self, parent):
        card_w = self.right_w - 40
        session_card = ctk.CTkFrame(
            parent, width=card_w, height=150,
            fg_color="#ffffff", corner_radius=10,
            border_width=2, border_color="#e74c3c"
        )
        session_card.place(x=20, y=400)

        ctk.CTkLabel(
            session_card, text="Current Session Info",
            font=self.F(18, True), text_color="#1a2a3a"
        ).place(x=card_w // 2, y=18, anchor="center")

        ctk.CTkLabel(session_card, text="พบตำหนิ: 19 จาน",
                     font=self.F(16, True), text_color="#e74c3c").place(x=40, y=50)

        ctk.CTkLabel(session_card, text="จานที่ : 81",
                     font=self.F(15), text_color="#1a2a3a").place(x=40, y=78)

        lot_date = datetime.now().strftime("%d%m%Y")
        ctk.CTkLabel(session_card, text=f"รหัสล็อต : PTP{lot_date}_01",
                     font=self.F(15), text_color="#1a2a3a").place(x=40, y=104)

    # Bottom
    def create_bottom_panel(self,):
        bottom_frame = ctk.CTkFrame(
            self.app, width=self.bottom_w, height=self.BOTTOM_H,
            fg_color="#ffffff", corner_radius=15,
            border_width=2, border_color="#aed6f1"
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
                                    fg_color="#1a2a3a", corner_radius=8)
        header_frame.place(x=5, y=5)

        col1_x = int((table_w - 10) * 0.32)
        col2_x = int((table_w - 10) * 0.76)

        ctk.CTkLabel(header_frame, text="Defect Type", font=self.F(16, True),
                     text_color="white").place(x=col1_x, y=22, anchor="center")
        ctk.CTkLabel(header_frame, text="Detection Status", font=self.F(16, True),
                     text_color="white").place(x=col2_x, y=22, anchor="center")

        row_start_y = 55
        row_h = 36
        for i, (defect, status, color) in enumerate(self.defect_data):
            row_y = row_start_y + i * row_h
            row_color = "#ffffff" if i % 2 == 0 else "#e0f2f7"
            row_frame = ctk.CTkFrame(table_frame, width=table_w - 10, height=row_h,
                                     fg_color=row_color, corner_radius=0)
            row_frame.place(x=5, y=row_y)

            ctk.CTkLabel(row_frame, text=defect, font=self.F(15),
                         text_color="#1a2a3a").place(x=col1_x, y=row_h // 2, anchor="center")

            status_color = "#199129" if color == "green" else "#e74c3c"
            ctk.CTkLabel(row_frame, text=status, font=self.F(15, True),
                         text_color=status_color).place(x=col2_x, y=row_h // 2, anchor="center")

    # ----------------- Export helpers (UTF-8 BOM + Excel Safe) -----------------
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
            w.writerow(["วันที่", "เวลา", "Plate ID", "Lot ID", "ตำหนิที่พบ", "หมายเหตุ"])
            for r in self.session_rows:
                defects_for_csv = self._excel_safe(r["defects"])
                w.writerow([r["date"], r["time"], r["plate_id"], r["lot_id"], defects_for_csv, self._excel_safe(r["note"])])
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

    # ---------- Export dialog ----------
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
            text_color="#FFFFFF", fg_color="#3498db", hover_color="#2980b9",
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

    def save_to_json(self, dialog):
        if not self.session_rows:
            messagebox.showwarning("Warning", "ยังไม่มีข้อมูลสำหรับบันทึก")
            return
        directory = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึก JSON")
        if directory:
            json_path = self._write_json(directory)
            messagebox.showinfo("Success", f"บันทึกไฟล์เรียบร้อย\n{os.path.basename(json_path)}")
            dialog.destroy()

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

    # ----------------- Camera update -----------------
    def start_camera(self):
        if self.cap:
            self.camera_running = True
            self.update_camera()

    def stop_camera(self):
        self.camera_running = False
        if self.cap:
            self.cap.release(); self.cap = None

    def update_camera(self):
        if not self.camera_running or not self.cap:
            return
        ret, frame = self.cap.read()
        if not ret:
            self.app.after(30, self.update_camera)
            return
        try:
            frame_resized = cv2.resize(frame, (self.cam_w, self.cam_h))
            cv2.rectangle(frame_resized, (20, 20),
                          (frame_resized.shape[1]-20, self.cam_h-20),
                          (0, 255, 0), 2)
            frame_rgb = cv2.cvtColor(frame_resized, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(frame_rgb)
            imgtk = ImageTk.PhotoImage(pil_img)
            self.camera_label.configure(image=imgtk, text="")
            self.camera_label.image = imgtk
        except Exception as e:
            print(f"Camera error: {e}")
        self.app.after(30, self.update_camera)

    # ----------------- Data collection loop -----------------
    def start_collect_loop(self):
        if not self.session_meta:
            self.session_meta = {
                "start_time": datetime.now().strftime("%H:%M:%S"),
                "lot_id": self.lot_id
            }
        self.session_rows.append(self.next_mock_row())
        self.collect_job_id = self.app.after(self.collect_interval_ms, self.start_collect_loop)

    def cancel_collect_loop(self):
        if self.collect_job_id is not None:
            try:
                self.app.after_cancel(self.collect_job_id)
            except Exception:
                pass
            self.collect_job_id = None

    # ----------------- Events -----------------
    def toggle_data_collection(self):
        if not self.is_collecting_data:
            self.is_collecting_data = True
            self.toggle_button.configure(text="หยุด", fg_color="#e74c3c", hover_color="#c0392b")
            # ถ้าอยากเริ่ม session ใหม่ทุกครั้ง:
            # self.session_rows = []; self.plate_id_counter, self.last_log_time = 1, None
            self.start_collect_loop()
        else:
            if messagebox.askyesno("ยืนยันการหยุด", "ท่านต้องการหยุดและบันทึกไฟล์หรือไม่?"):
                self.is_collecting_data = False
                self.toggle_button.configure(text="เริ่ม", fg_color="#3498db", hover_color="#2980b9")
                self.cancel_collect_loop()
                directory = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึกไฟล์รายงาน")
                if directory:
                    self.export_session_csv_and_json(directory)
            else:
                return

    def create_mock_data(self):
        self.mock_data = {
            "session_start": datetime.now().strftime("%H:%M:%S"),
            "plates_processed": 81,
            "defects_found": 19
        }

    def update_header_time(self):
        self.header_time_label.configure(text=datetime.now().strftime("%H:%M:%S"))
        self.app.after(1000, self.update_header_time)

    def on_closing(self):
        if self.is_collecting_data:
            if not messagebox.askyesno("ออกจากโปรแกรม", "กำลังบันทึกข้อมูลอยู่ ต้องการออกทันทีหรือไม่?"):
                return
        self.cancel_collect_loop()
        self.stop_camera()
        self.app.destroy()

    def run(self):
        self.app.mainloop()


if __name__ == "__main__":
    app = LeafPlateDetectionApp()
    app.run()
