import customtkinter as ctk
import tkinter as tk
import cv2
from PIL import Image, ImageTk
from datetime import datetime
import os
import threading
import json


class QualityControlApp:
    def __init__(self):
        self.setup_app()
        self.setup_camera()
        self.create_widgets()
        self.create_mock_data()
        self.start_camera()

    def setup_app(self):
        """Initialize the main application window"""
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.app = ctk.CTk()
        self.app.title("Leaf Plate Defect Detection")
        self.app.geometry("1200x800")
        self.app.configure(fg_color="white")

        # Initialize shape counts BEFORE creating widgets
        self.shape_counts = {
            "heart": 28,
            "rectangle": 31,
            "oval": 22,
            "total": 81
        }

        # Defect data matching the screenshot exactly
        self.defect_data = [
            ("รอยแตก", "ผ่าน"),
            ("รอยยุบหรือรอยพับ", "พบตำหนิ"),
            ("จุดใหญ่หรือสีคล้ำ", "ผ่าน"),
            ("รูพรุนหรือรูเย็บ", "ผ่าน"),
            ("รอยยีดวนหรือรอยกลอก", "ผ่าน")
        ]

        # Handle app closing
        self.app.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_camera(self):
        """Initialize camera variables"""
        self.cap = cv2.VideoCapture(0, cv2.CAP_AVFOUNDATION)
        if not self.cap.isOpened():
            print("Cannot open camera at index 0 with CAP_AVFOUNDATION")
            self.cap = None
        else:
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

        self.current_frame = None
        self.camera_running = False

    def start_camera(self):
        """Start the camera frame update loop"""
        if self.cap:
            self.camera_running = True
            self.update_camera()

    def stop_camera(self):
        """Stop the camera and release resources"""
        self.camera_running = False
        if self.cap:
            self.cap.release()
            self.cap = None

    def update_camera(self):
        """Capture frame and update GUI, scheduled with after"""
        if not self.camera_running or not self.cap:
            return

        ret, frame = self.cap.read()
        if not ret:
            print("Failed to read frame")
            self.app.after(30, self.update_camera)
            return

        try:
            # Resize frame to 800x400 to fill camera_frame
            frame_resized = cv2.resize(frame, (800, 400))

            # Draw rectangle box on the resized frame (green box, thickness 2)
            box_color = (0, 255, 0)  # Green in BGR
            box_thickness = 2
            height, width, _ = frame_resized.shape
            margin = 20
            top_left = (margin, margin)
            bottom_right = (width - margin, height - margin)
            cv2.rectangle(frame_resized, top_left, bottom_right, box_color, box_thickness)

            frame_rgb = cv2.cvtColor(frame_resized, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(frame_rgb)
            self.current_frame = pil_img

            imgtk = ImageTk.PhotoImage(self.current_frame)
            self.camera_label.configure(image=imgtk, text="")
            self.camera_label.image = imgtk  # Keep reference to avoid GC

        except Exception as e:
            print(f"Error processing frame: {e}")

        # Schedule next frame update with shorter interval (~30 FPS)
        self.app.after(30, self.update_camera)

    def on_closing(self):
        """Handle application closing"""
        self.stop_camera()
        self.app.destroy()

    def create_mock_data(self):
        """Create mock data for JSON export"""
        current_date = datetime.now()
        lot_date_format = current_date.strftime("%d%m%Y")
        display_date_format = current_date.strftime("%d/%m/%Y")

        self.mock_json_data = {
            "leaf_plate_reports": {
                f"PTP{lot_date_format}_01": {
                    "date": display_date_format,
                    "shape_counts": self.shape_counts,
                    "plates": {
                        "1": {"time": "18:53:00", "defects": ["รอยย่น", "จุดไหม้"], "shape": "heart"},
                        "2": {"time": "18:54:00", "defects": [], "shape": "rectangle"},
                        "3": {"time": "18:55:00", "defects": ["รอยแตก"], "shape": "oval"},
                        "4": {"time": "18:56:00", "defects": [], "shape": "heart"},
                        "5": {"time": "18:57:00", "defects": ["รอยขีดข่วน", "รูเข็ม"], "shape": "rectangle"},
                        "6": {"time": "18:58:00", "defects": [], "shape": "oval"},
                        "7": {"time": "18:59:00", "defects": ["รอยย่น"], "shape": "heart"},
                        "8": {"time": "19:00:00", "defects": [], "shape": "rectangle"},
                        "9": {"time": "19:01:00", "defects": [], "shape": "oval"},
                        "10": {"time": "19:02:00", "defects": ["จุดไหม้"], "shape": "heart"}
                    }
                }
            }
        }

    def create_widgets(self):
        """Create all widgets matching the screenshot layout exactly"""
        # Title - positioned exactly like in screenshot
        title_label = ctk.CTkLabel(
            self.app,
            text="Leaf Plate Defect Detection",
            font=("Arial", 18),
            text_color="#999999",
            fg_color="transparent"
        )
        title_label.place(x=25, y=15)

        # Main content area
        self.create_main_layout()

    def create_main_layout(self):
        """Create the main layout matching screenshot proportions"""
        # Camera area frame
        self.camera_frame = ctk.CTkFrame(
            self.app,
            width=800,
            height=400,
            fg_color="#CCCCCC",
            corner_radius=0,
            border_width=1,
            border_color="#999999"
        )
        self.camera_frame.place(x=25, y=70)

        # Camera label for displaying video - standard tkinter Label for Mac M1 compatibility
        self.camera_label = tk.Label(
            self.camera_frame,
            text="Initializing Camera...",
            font=("Arial", 20),
            fg="black",
            bg="#CCCCCC"
        )
        # Adjust camera_label size to match resized frame 800x400
        self.camera_label.place(x=1, y=1, width=800, height=400)

        # Right panel - Purple section
        self.create_purple_section()

        # Right panel - Shape counting section
        self.create_shape_counting_section()

        # Right panel - Gray info section
        self.create_gray_info_section()

        # Green export button
        self.create_export_button()

        # Bottom table
        self.create_bottom_table()

    def create_purple_section(self):
        """Create the purple count section exactly as in screenshot"""
        purple_frame = ctk.CTkFrame(
            self.app,
            width=350,
            height=180,
            fg_color="#E6E6FA",
            corner_radius=0,
            border_width=1,
            border_color="#999999"
        )
        purple_frame.place(x=840, y=70)

        # Thai text
        thai_label = ctk.CTkLabel(
            purple_frame,
            text="จำนวนที่นับได้",
            font=("Arial", 24, "bold"),
            text_color="black",
            fg_color="transparent"
        )
        thai_label.place(x=175, y=40, anchor="center")

        # Large number 81
        number_label = ctk.CTkLabel(
            purple_frame,
            text=str(self.shape_counts["total"]),
            font=("Arial", 80, "bold"),
            text_color="black",
            fg_color="transparent"
        )
        number_label.place(x=175, y=120, anchor="center")

    def create_shape_counting_section(self):
        """Create a new section for shape counting"""
        shape_frame = ctk.CTkFrame(
            self.app,
            width=350,
            height=120,
            fg_color="#F0F8FF",
            corner_radius=0,
            border_width=1,
            border_color="#999999"
        )
        shape_frame.place(x=840, y=250)

        # Title
        title_label = ctk.CTkLabel(
            shape_frame,
            text="จำนวนตามรูปแบบ",
            font=("Arial", 16, "bold"),
            text_color="black",
            fg_color="transparent"
        )
        title_label.place(x=175, y=15, anchor="center")

        # Shape counts in a grid layout
        # Row 1: Heart and Rectangle
        heart_label = ctk.CTkLabel(
            shape_frame,
            text="♥ หัวใจ:",
            font=("Arial", 14, "bold"),
            text_color="#FF1493",
            fg_color="transparent"
        )
        heart_label.place(x=20, y=45)

        heart_count = ctk.CTkLabel(
            shape_frame,
            text=str(self.shape_counts["heart"]),
            font=("Arial", 14, "bold"),
            text_color="black",
            fg_color="transparent"
        )
        heart_count.place(x=90, y=45)

        rectangle_label = ctk.CTkLabel(
            shape_frame,
            text="▬ สี่เหลี่ยม:",
            font=("Arial", 14, "bold"),
            text_color="#4169E1",
            fg_color="transparent"
        )
        rectangle_label.place(x=180, y=45)

        rectangle_count = ctk.CTkLabel(
            shape_frame,
            text=str(self.shape_counts["rectangle"]),
            font=("Arial", 14, "bold"),
            text_color="black",
            fg_color="transparent"
        )
        rectangle_count.place(x=280, y=45)

        # Row 2: Oval
        oval_label = ctk.CTkLabel(
            shape_frame,
            text="⬭ วงรี:",
            font=("Arial", 14, "bold"),
            text_color="#32CD32",
            fg_color="transparent"
        )
        oval_label.place(x=20, y=80)

        oval_count = ctk.CTkLabel(
            shape_frame,
            text=str(self.shape_counts["oval"]),
            font=("Arial", 14, "bold"),
            text_color="black",
            fg_color="transparent"
        )
        oval_count.place(x=90, y=80)

    def create_gray_info_section(self):
        """Create the gray information section exactly as in screenshot"""
        gray_frame = ctk.CTkFrame(
            self.app,
            width=350,
            height=180,
            fg_color="#CCCCCC",
            corner_radius=0,
            border_width=1,
            border_color="#999999"
        )
        gray_frame.place(x=840, y=370)

        # Red defect text
        defect_label = ctk.CTkLabel(
            gray_frame,
            text="มีตำหนิ 19 ใบ",
            font=("Arial", 18, "bold"),
            text_color="red",
            fg_color="transparent"
        )
        defect_label.place(x=25, y=20)

        # Plate ID
        plate_label = ctk.CTkLabel(
            gray_frame,
            text="Plate ID : 81",
            font=("Arial", 14),
            text_color="black",
            fg_color="transparent"
        )
        plate_label.place(x=25, y=70)

        # Lot ID with current date
        current_date = datetime.now()
        lot_date_format = current_date.strftime("%d%m%Y")
        lot_label = ctk.CTkLabel(
            gray_frame,
            text=f"Lot ID : PTP{lot_date_format}_01",
            font=("Arial", 14),
            text_color="black",
            fg_color="transparent"
        )
        lot_label.place(x=25, y=95)

        # Time (blue text)
        current_time_str = datetime.now().strftime("%H:%M:%S")
        self.time_label = ctk.CTkLabel(
            gray_frame,
            text=current_time_str,
            font=("Arial", 14, "bold"),
            text_color="blue",
            fg_color="transparent"
        )
        self.time_label.place(x=25, y=140)

        # Date (blue text) - Thai Buddhist calendar
        year = datetime.now().year + 543
        thai_date = datetime.now().strftime(f"%d/%m/{year}")
        date_label = ctk.CTkLabel(
            gray_frame,
            text=thai_date,
            font=("Arial", 14, "bold"),
            text_color="blue",
            fg_color="transparent"
        )
        date_label.place(x=250, y=140)

        # Start updating the time label every second
        self.update_time_label()

    def create_export_button(self):
        """Create the green export button exactly as in screenshot"""
        export_button = ctk.CTkButton(
            self.app,
            width=200,
            height=60,
            text="ส่งออก",
            font=("Arial", 20, "bold"),
            fg_color="#90EE90",
            text_color="black",
            hover_color="#7CFC00",
            corner_radius=5,
            border_width=1,
            border_color="#999999",
            command=self.on_export_click
        )
        export_button.place(x=990, y=710)

        # Start/Stop toggle button
        self.toggle_button = ctk.CTkButton(
            self.app,
            width=150,
            height=50,
            text="เริ่ม",
            font=("Arial", 16, "bold"),
            fg_color="#32CD32",
            text_color="white",
            hover_color="#228B22",
            corner_radius=5,
            border_width=1,
            border_color="#999999",
            command=self.toggle_data_collection
        )
        self.toggle_button.place(x=840, y=710)

    def toggle_data_collection(self):
        """Toggle between start and stop data collection"""
        if getattr(self, "is_collecting_data", False):
            # Currently collecting, so stop and save
            self.is_collecting_data = False
            self.toggle_button.configure(text="เริ่ม", fg_color="#32CD32", hover_color="#228B22")
            self.save_collected_to_csv_and_json()
        else:
            # Start collecting
            self.is_collecting_data = True
            self.collected_data = {
                "leaf_plate_reports": {}
            }
            self.collection_start_time = datetime.now()
            self.toggle_button.configure(text="ส่งออก", fg_color="#FF4500", hover_color="#B22222")
            self.collect_data_periodically()

    def collect_data_periodically(self):
        """Collect data periodically while collection is active"""
        if not getattr(self, "is_collecting_data", False):
            return

        # Simulate data collection: add a new plate entry with current time
        current_time = datetime.now().strftime("%H:%M:%S")
        current_date = datetime.now()
        lot_date_format = current_date.strftime("%d%m%Y")
        lot_id = f"PTP{lot_date_format}_01"

        if lot_id not in self.collected_data["leaf_plate_reports"]:
            self.collected_data["leaf_plate_reports"][lot_id] = {
                "date": current_date.strftime("%d/%m/%Y"),
                "shape_counts": self.shape_counts.copy(),
                "plates": {}
            }

        plates = self.collected_data["leaf_plate_reports"][lot_id]["plates"]
        new_plate_id = str(len(plates) + 1)
        # For demonstration, add a dummy defect and shape randomly or fixed
        new_plate_data = {
            "time": current_time,
            "defects": ["รอยแตก"],  # Example defect
            "shape": "heart"  # Example shape
        }
        plates[new_plate_id] = new_plate_data

        # Schedule next collection after 1 second
        self.app.after(1000, self.collect_data_periodically)

    def save_collected_to_csv_and_json(self):
        import csv
        from tkinter import messagebox

        try:
            save_directory = "/Users/rocket/Downloads/For test project/GUI"
            # Save CSV
            csv_filename = self.get_unique_filename(save_directory, ".csv")
            with open(csv_filename, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["Plate ID", "Time", "Defects", "Shape"])
                plates = \
                self.collected_data["leaf_plate_reports"][list(self.collected_data["leaf_plate_reports"].keys())[0]][
                    "plates"]
                for plate_id, plate_data in plates.items():
                    defects = ", ".join(plate_data["defects"])
                    writer.writerow([plate_id, plate_data["time"], defects, plate_data["shape"]])

            # Save JSON
            json_filename = self.get_unique_filename(save_directory, ".json")
            with open(json_filename, 'w', encoding='utf-8') as file:
                json.dump(self.collected_data, file, ensure_ascii=False, indent=2)

            messagebox.showinfo("Success",
                                f"CSV and JSON files saved successfully!\n{os.path.basename(csv_filename)}\n{os.path.basename(json_filename)}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save files:\n{str(e)}")

    def update_time_label(self):
        """Update the time label text every second"""
        current_time_str = datetime.now().strftime("%H:%M:%S")
        self.time_label.configure(text=current_time_str)
        self.app.after(1000, self.update_time_label)

    def create_bottom_table(self):
        """Create the bottom table exactly matching the screenshot"""
        # Table starts at y=490, spans the width of the left area
        table_y = 490
        table_width = 800
        row_height = 45

        # Table header
        header_frame = ctk.CTkFrame(
            self.app,
            width=table_width,
            height=row_height,
            fg_color="white",
            corner_radius=0,
            border_width=1,
            border_color="#999999"
        )
        header_frame.place(x=25, y=table_y)

        # Header labels
        ctk.CTkLabel(
            header_frame,
            text="ตำหนิ",
            font=("Arial", 16, "bold"),
            text_color="black",
            fg_color="transparent"
        ).place(x=240, y=22, anchor="center")

        ctk.CTkLabel(
            header_frame,
            text="สถานะ",
            font=("Arial", 16, "bold"),
            text_color="black",
            fg_color="transparent"
        ).place(x=720, y=22, anchor="center")

        # Table rows
        for i, (defect, status) in enumerate(self.defect_data):
            row_y = table_y + row_height + (i * row_height)

            row_frame = ctk.CTkFrame(
                self.app,
                width=table_width,
                height=row_height,
                fg_color="white",
                corner_radius=0,
                border_width=1,
                border_color="#999999"
            )
            row_frame.place(x=25, y=row_y)

            # Defect name
            ctk.CTkLabel(
                row_frame,
                text=defect,
                font=("Arial", 14),
                text_color="black",
                fg_color="transparent"
            ).place(x=240, y=22, anchor="center")

            # Status with color
            color = "red" if status == "พบตำหนิ" else "green"
            ctk.CTkLabel(
                row_frame,
                text=status,
                font=("Arial", 14, "bold"),
                text_color=color,
                fg_color="transparent"
            ).place(x=720, y=22, anchor="center")

    def update_shape_counts(self, heart=None, rectangle=None, oval=None):
        """Method to update shape counts dynamically"""
        if heart is not None:
            self.shape_counts["heart"] = heart
        if rectangle is not None:
            self.shape_counts["rectangle"] = rectangle
        if oval is not None:
            self.shape_counts["oval"] = oval

        # Update total
        self.shape_counts["total"] = (
                self.shape_counts["heart"] +
                self.shape_counts["rectangle"] +
                self.shape_counts["oval"]
        )

    def on_export_click(self):
        """Handle export button click"""
        self.show_export_options()

    def show_export_options(self):
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("Export Options")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self.app)
        dialog.grab_set()

        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (200 // 2)
        dialog.geometry(f"400x200+{x}+{y}")

        title_label = ctk.CTkLabel(
            dialog,
            text="เลือกการส่งออกข้อมูล",
            font=("Arial", 20, "bold"),
            text_color="black"
        )
        title_label.pack(pady=30)

        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=20)

        csv_button = ctk.CTkButton(
            button_frame,
            text="Save to CSV",
            fg_color="#4CAF50",
            text_color="white",
            font=("Arial", 16, "bold"),
            width=150,
            height=50,
            command=lambda: self.save_to_csv(dialog)
        )
        csv_button.pack(side="left", padx=10)

        json_button = ctk.CTkButton(
            button_frame,
            text="Save to JSON",
            fg_color="#FF9800",
            text_color="white",
            font=("Arial", 16, "bold"),
            width=150,
            height=50,
            command=lambda: self.save_to_json(dialog)
        )
        json_button.pack(side="right", padx=10)

    def get_unique_filename(self, base_path, extension):
        """Generate a unique filename in the base_path with the given extension"""
        current_date = datetime.now().strftime("%d%m%Y")
        counter = 1
        while True:
            filename = f"Report_{current_date}_{counter}{extension}"
            full_path = os.path.join(base_path, filename)
            if not os.path.exists(full_path):
                return full_path
            counter += 1

    def save_to_csv(self, dialog):
        import csv
        from tkinter import filedialog, messagebox

        try:
            save_directory = filedialog.askdirectory(title="Select Directory to Save CSV Report")
            if save_directory:
                filename = self.get_unique_filename(save_directory, ".csv")

                with open(filename, 'w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    # Write header
                    writer.writerow(["Plate ID", "Time", "Defects", "Shape"])
                    # Write data rows
                    plates = self.mock_json_data["leaf_plate_reports"][
                        list(self.mock_json_data["leaf_plate_reports"].keys())[0]]["plates"]
                    for plate_id, plate_data in plates.items():
                        defects = ", ".join(plate_data["defects"])
                        writer.writerow([plate_id, plate_data["time"], defects, plate_data["shape"]])

                messagebox.showinfo("Success", f"CSV file saved successfully!\n{os.path.basename(filename)}")
                dialog.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV file:\n{str(e)}")

    def save_to_json(self, dialog):
        from tkinter import filedialog, messagebox

        try:
            save_directory = filedialog.askdirectory(title="Select Directory to Save JSON Report")
            if save_directory:
                filename = self.get_unique_filename(save_directory, ".json")

                with open(filename, 'w', encoding='utf-8') as file:
                    json.dump(self.mock_json_data, file, ensure_ascii=False, indent=2)

                messagebox.showinfo("Success", f"JSON file saved successfully!\n{os.path.basename(filename)}")
                dialog.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save JSON file:\n{str(e)}")

    def export_csv(self, dialog):
        """Export mock data to CSV"""
        import csv
        try:
            from tkinter import filedialog, messagebox

            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if file_path:
                with open(file_path, "w", encoding="utf-8", newline='') as f:
                    writer = csv.writer(f)
                    # Write header row like the screenshot
                    writer.writerow(["Plate ID", "Time", "Defects", "Shape"])

                    plates = self.mock_json_data["leaf_plate_reports"][
                        list(self.mock_json_data["leaf_plate_reports"].keys())[0]]["plates"]

                    for plate_id, data in plates.items():
                        defects = ", ".join(data["defects"]) if data["defects"] else "None"
                        writer.writerow([plate_id, data["time"], defects, data["shape"]])

                messagebox.showinfo("Export CSV", f"Data exported successfully to:\n{file_path}")
                dialog.destroy()
        except Exception as e:
            print(f"Export CSV error: {e}")

    def run(self):
        self.app.mainloop()


if __name__ == "__main__":
    app = QualityControlApp()
    app.run()
