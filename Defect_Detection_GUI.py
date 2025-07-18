import customtkinter as ctk
import tkinter as tk
import cv2
from PIL import Image, ImageTk
from datetime import datetime
import os
import threading
import json
import csv
from tkinter import filedialog, messagebox


class LeafPlateDetectionApp:
    """
    Redesigned Leaf Plate Defect Detection Application
    Improved layout with better organization and user experience
    """

    def __init__(self):
        self.initialize_data()
        self.setup_app()
        self.setup_camera()
        self.create_widgets()
        self.create_mock_data()
        self.start_camera()

    def initialize_data(self):
        """Initialize all data structures and variables"""
        self.shape_counts = {
            "heart": 28,
            "rectangle": 31,
            "circle": 10,
            "oval": 22,
            "total": 81
        }
        self.defect_data = [
            ("รอยแตก", "ผ่าน", "green"),
            ("รอยยุบหรือรอยพับ", "พบตำหนิ", "red"),
            ("จุดใหญ่หรือสีคล้ำ", "ผ่าน", "green"),
            ("รูพรุนหรือรูเข็บ", "ผ่าน", "green"),
            ("รอยขีดข่วนหรือรอยกลอก", "ผ่าน", "green")
        ]
        self.is_collecting_data = False
        self.collected_data = {"leaf_plate_reports": {}}
        self.collection_start_time = None
        self.current_frame = None
        self.camera_running = False
        self.cap = None
        # Mock data for export
        self.mock_json_data = {
            "leaf_plate_reports": {
                "session_1": {
                    "plates": {
                        "plate_001": {
                            "time": "14:30:15",
                            "defects": ["รอยแตก"],
                            "shape": "heart"
                        },
                        "plate_002": {
                            "time": "14:31:20",
                            "defects": [],
                            "shape": "rectangle"
                        },
                        "plate_003": {
                            "time": "14:32:25",
                            "defects": ["รอยยุบ"],
                            "shape": "oval"
                        }
                    }
                }
            }
        }

    def setup_app(self):
        """Initialize the main application window"""
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")  # Keep default blue theme for base components
        self.app = ctk.CTk()
        self.app.title("Leaf Plate Defect Detection - Redesigned")
        self.app.geometry("1400x900")
        # Main app background: White (30%)
        self.app.configure(fg_color="#ffffff")
        self.app.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_camera(self):
        """Initialize camera system"""
        self.cap = cv2.VideoCapture(0, cv2.CAP_AVFOUNDATION)
        if not self.cap.isOpened():
            print("Cannot open camera")
            self.cap = None
        else:
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

    def create_widgets(self):
        """Create the redesigned user interface"""
        self.create_header()
        self.create_main_content()

    def create_header(self):
        """Create application header with title and time"""
        header_frame = ctk.CTkFrame(
            self.app,
            width=1350,
            height=80,
            # Header background: Dark Blue (60%)
            fg_color="#1a2a3a",
            corner_radius=10
        )
        header_frame.place(x=25, y=20)

        # Title text: White
        title_label = ctk.CTkLabel(
            header_frame,
            text="Leaf Plate Defect Detection System",
            font=("Arial", 28, "bold"),
            text_color="white"
        )
        title_label.place(x=50, y=25)

        # Time display text: White
        self.header_time_label = ctk.CTkLabel(
            header_frame,
            text=datetime.now().strftime("%H:%M:%S"),
            font=("Arial", 20, "bold"),
            text_color="white"
        )
        self.header_time_label.place(x=1100, y=15)

        # Date display text: White
        current_date = datetime.now()
        thai_year = current_date.year + 543
        date_str = current_date.strftime(f"%d/%m/{thai_year}")
        self.header_date_label = ctk.CTkLabel(
            header_frame,
            text=date_str,
            font=("Arial", 16),
            text_color="white"
        )
        self.header_date_label.place(x=1100, y=45)

        # Start time updates
        self.update_header_time()

    def create_main_content(self):
        """Create main content area with improved layout"""
        # Left panel - Camera and controls
        self.create_left_panel()
        # Right panel - Statistics and info
        self.create_right_panel()
        # Bottom panel - Defect table
        self.create_bottom_panel()

    def create_left_panel(self):
        """Create left panel with camera and controls"""
        left_frame = ctk.CTkFrame(
            self.app,
            width=680,
            height=500,
            fg_color="#ffffff",  # White (30%)
            corner_radius=15,
            border_width=2,
            border_color="#aed6f1"  # Light Blue border
        )
        left_frame.place(x=25, y=120)

        # Camera display frame: Light Blue (60%)
        self.camera_frame = ctk.CTkFrame(
            left_frame,
            width=640,
            height=400,
            fg_color="#e0f2f7",
            corner_radius=10,
            border_width=1,
            border_color="#aed6f1"  # Light Blue border
        )
        self.camera_frame.place(x=20, y=20)
        self.camera_label = tk.Label(
            self.camera_frame,
            text="Initializing Camera...",
            font=("Arial", 16),
            fg="black",  # Black text on light background
            bg="#e0f2f7"  # Light Blue background
        )
        self.camera_label.place(x=1, y=1, width=640, height=400)

        # Control buttons
        self.create_control_buttons(left_frame)

    def create_control_buttons(self, parent):
        """Create control buttons in the left panel"""
        button_frame = ctk.CTkFrame(
            parent,
            width=640,
            height=50,
            fg_color="transparent"
        )
        button_frame.place(x=20, y=430)

        # Start/Stop button: Red for Stop, Medium Blue for Start
        self.toggle_button = ctk.CTkButton(
            button_frame,
            width=120,
            height=40,
            text="เริ่ม",
            font=("Arial", 14, "bold"),
            text_color="#FFFFFF",
            fg_color="#3498db",  # Medium Blue for "Start"
            hover_color="#FFD700",  # Yellow for warning
            command=self.toggle_data_collection
        )
        self.toggle_button.place(x=10, y=5)

        # Export button: Medium Blue (60%)
        export_button = ctk.CTkButton(
            button_frame,
            width=120,
            height=40,
            text="Export",
            font=("Arial", 14, "bold"),
            text_color="#FFFFFF",
            fg_color="#50C878",  # Emerald Green
            hover_color="#27ad60",  # Dark Green
            command=self.show_export_dialog
        )
        export_button.place(x=510, y=5)

    def create_right_panel(self):
        """Create right panel with statistics"""
        right_frame = ctk.CTkFrame(
            self.app,
            width=645,
            height=500,
            fg_color="#ffffff",  # White (30%)
            corner_radius=15,
            border_width=2,
            border_color="#aed6f1"  # Light Blue border
        )
        right_frame.place(x=730, y=120)

        # Total count card
        self.create_total_count_card(right_frame)
        # Shape counts card
        self.create_shape_counts_card(right_frame)
        # Session info card
        self.create_session_info_card(right_frame)

    def create_total_count_card(self, parent):
        """Create total count display card"""
        total_card = ctk.CTkFrame(
            parent,
            width=600,
            height=120,
            fg_color="#e0f2f7",  # Light Blue (60%)
            corner_radius=10,
            border_width=2,
            border_color="#3498db"  # Medium Blue border
        )
        total_card.place(x=20, y=20)

        # Total count label text: Dark Blue
        total_label = ctk.CTkLabel(
            total_card,
            text="Total Plates Counted",
            font=("Arial", 16, "bold"),
            text_color="#1a2a3a"
        )
        total_label.place(x=300, y=20, anchor="center")

        # Total number text: Medium Blue
        self.total_number_label = ctk.CTkLabel(
            total_card,
            text=str(self.shape_counts["total"]),
            font=("Arial", 48, "bold"),
            text_color="#e74c3c"
        )
        self.total_number_label.place(x=300, y=70, anchor="center")

    def create_shape_counts_card(self, parent):
        """Create shape counts display card"""
        shape_card = ctk.CTkFrame(
            parent,
            width=600,
            height=140,
            fg_color="#e0f2f7",  # Light Blue (60%)
            corner_radius=10,
            border_width=2,
            border_color="#3498db"  # Medium Blue border
        )
        shape_card.place(x=20, y=150)

        # Shape counts title text: Dark Blue
        shape_title = ctk.CTkLabel(
            shape_card,
            text="จานแต่ละรูปแบบ",
            font=("Arial", 16, "bold"),
            text_color="#1a2a3a"
        )
        shape_title.place(x=300, y=15, anchor="center")

        # Heart count frame: White (30%), text Red (10%)
        heart_frame = ctk.CTkFrame(shape_card, width=180, height=80, fg_color="#ffffff", corner_radius=8)
        heart_frame.place(x=20, y=45)
        ctk.CTkLabel(heart_frame, text="Heart", font=("Arial", 12, "bold"), text_color="#e74c3c").place(x=90, y=25,
                                                                                                        anchor="center")
        ctk.CTkLabel(heart_frame, text=str(self.shape_counts["heart"]), font=("Arial", 18, "bold"),
                     text_color="#1a2a3a").place(x=90, y=55, anchor="center")

        # Rectangle count frame: White (30%), text Medium Blue (60%)
        rect_frame = ctk.CTkFrame(shape_card, width=180, height=80, fg_color="#ffffff", corner_radius=8)
        rect_frame.place(x=210, y=45)
        ctk.CTkLabel(rect_frame, text="Rectangle", font=("Arial", 12, "bold"), text_color="#3498db").place(x=90, y=25,
                                                                                                           anchor="center")
        ctk.CTkLabel(rect_frame, text=str(self.shape_counts["rectangle"]), font=("Arial", 18, "bold"),
                     text_color="#1a2a3a").place(x=90, y=55, anchor="center")

        # Oval count frame: White (30%), text a slightly darker blue (60%)
        oval_frame = ctk.CTkFrame(shape_card, width=180, height=80, fg_color="#ffffff", corner_radius=8)
        oval_frame.place(x=400, y=45)
        ctk.CTkLabel(oval_frame, text="Oval", font=("Arial", 12, "bold"), text_color="#2c7bb6").place(x=90, y=25,
                                                                                                      anchor="center")
        ctk.CTkLabel(oval_frame, text=str(self.shape_counts["oval"]), font=("Arial", 18, "bold"),
                     text_color="#1a2a3a").place(x=90, y=55, anchor="center")

    def create_session_info_card(self, parent):
        """Create session information card"""
        session_card = ctk.CTkFrame(
            parent,
            width=600,
            height=140,
            fg_color="#ffffff",  # White (30%)
            corner_radius=10,
            border_width=2,
            border_color="#e74c3c"  # Red border (10%)
        )
        session_card.place(x=20, y=300)

        # Session info title text: Dark Blue
        session_title = ctk.CTkLabel(
            session_card,
            text="Current Session Info",
            font=("Arial", 16, "bold"),
            text_color="#1a2a3a"
        )
        session_title.place(x=300, y=15, anchor="center")

        # Defect alert text: Red (10%) ตรวจพบตำหนิ
        defect_label = ctk.CTkLabel(
            session_card,
            text="Defects Found: 19 plates",
            font=("Arial", 16, "bold"),
            text_color="#e74c3c"
        )
        defect_label.place(x=50, y=50)

        # Plate ID text: Dark Blue
        plate_label = ctk.CTkLabel(
            session_card,
            text="Current Plate ID: 81",
            font=("Arial", 14),
            text_color="#1a2a3a"
        )
        plate_label.place(x=50, y=80)

        # Lot ID text: Dark Blue
        current_date = datetime.now()
        lot_date = current_date.strftime("%d%m%Y")
        lot_label = ctk.CTkLabel(
            session_card,
            text=f"Lot ID: PTP{lot_date}_01",
            font=("Arial", 14),
            text_color="#1a2a3a"
        )
        lot_label.place(x=50, y=105)

    def create_bottom_panel(self):
        """Create bottom panel with defect status table"""
        bottom_frame = ctk.CTkFrame(
            self.app,
            width=1350,
            height=220,
            fg_color="#ffffff",  # White (30%)
            corner_radius=15,
            border_width=2,
            border_color="#aed6f1"  # Light Blue border
        )
        bottom_frame.place(x=25, y=640)

        # Create table
        self.create_defect_table(bottom_frame)

    def create_defect_table(self, parent):
        """Create defect status table with improved design"""
        table_frame = ctk.CTkFrame(
            parent,
            width=1310,
            height=190,
            fg_color="#ffffff",  # White (30%)
            corner_radius=10
        )
        table_frame.place(x=20, y=20)

        # Table header: Dark Blue (60%)
        header_frame = ctk.CTkFrame(
            table_frame,
            width=1300,
            height=40,
            fg_color="#1a2a3a",
            corner_radius=8
        )
        header_frame.place(x=5, y=5)

        # Adjusted positions for remaining columns after removing "Action"
        ctk.CTkLabel(
            header_frame,
            text="Defect Type",
            font=("Arial", 14, "bold"),
            text_color="white"
        ).place(x=350, y=20, anchor="center")  # Shifted right
        ctk.CTkLabel(
            header_frame,
            text="Detection Status",
            font=("Arial", 14, "bold"),
            text_color="white"
        ).place(x=950, y=20, anchor="center")  # Shifted right

        # Table rows
        for i, (defect, status, color) in enumerate(self.defect_data):
            row_y = 55 + (i * 25)
            # Alternating row colors: White and Light Blue (30% white, 60% blue)
            row_color = "#ffffff" if i % 2 == 0 else "#e0f2f7"
            row_frame = ctk.CTkFrame(
                table_frame,
                width=1300,
                height=25,
                fg_color=row_color,
                corner_radius=0
            )
            row_frame.place(x=5, y=row_y)

            # Defect name text: Dark Blue
            ctk.CTkLabel(
                row_frame,
                text=defect,
                font=("Arial", 12),
                text_color="#1a2a3a"
            ).place(x=350, y=12, anchor="center")  # Shifted right

            # Status with color coding: Medium Blue for "Pass", Red for "Defect" (60% blue, 10% red)
            status_color = "#3498db" if color == "green" else "#e74c3c"
            status_text = f"{status}"
            ctk.CTkLabel(
                row_frame,
                text=status_text,
                font=("Arial", 12, "bold"),
                text_color=status_color
            ).place(x=950, y=12, anchor="center")  # Shifted right

            # Removed Action button as requested

    # Export functionality
    def get_unique_filename(self, base_path, extension):
        """Generate unique filename to avoid overwrites"""
        current_date = datetime.now().strftime("%d%m%Y")
        counter = 1
        while True:
            filename = f"Report_{current_date}_{counter}{extension}"
            full_path = os.path.join(base_path, filename)
            if not os.path.exists(full_path):
                return full_path
            counter += 1

    def save_to_csv(self, dialog):
        """Export data to CSV format"""
        try:
            save_directory = filedialog.askdirectory(
                title="Select Directory to Save CSV Report"
            )
            if save_directory:
                filename = self.get_unique_filename(save_directory, ".csv")
                with open(filename, 'w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(["Plate ID", "Time", "Defects", "Shape"])
                    # Extract plate data from mock data
                    plates = self.mock_json_data["leaf_plate_reports"][
                        list(self.mock_json_data["leaf_plate_reports"].keys())[0]
                    ]["plates"]
                    for plate_id, plate_data in plates.items():
                        defects = ", ".join(plate_data["defects"])
                        writer.writerow([
                            plate_id,
                            plate_data["time"],
                            defects,
                            plate_data["shape"]
                        ])
                messagebox.showinfo(
                    "Success",
                    f"CSV file saved successfully!\n{os.path.basename(filename)}"
                )
                dialog.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV file:\n{str(e)}")

    def save_to_json(self, dialog):
        """Export data to JSON format"""
        try:
            save_directory = filedialog.askdirectory(
                title="Select Directory to Save JSON Report"
            )
            if save_directory:
                filename = self.get_unique_filename(save_directory, ".json")
                with open(filename, 'w', encoding='utf-8') as file:
                    json.dump(self.mock_json_data, file, ensure_ascii=False, indent=2)
                messagebox.showinfo(
                    "Success",
                    f"JSON file saved successfully!\n{os.path.basename(filename)}"
                )
                dialog.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save JSON file:\n{str(e)}")

    def save_collected_to_csv_and_json(self, dialog):
        """Save collected real-time data to both CSV and JSON"""
        try:
            save_directory = filedialog.askdirectory(
                title="Select Directory to Save Both Files"
            )
            if save_directory:
                # Save CSV
                csv_filename = self.get_unique_filename(save_directory, ".csv")
                with open(csv_filename, 'w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(["Plate ID", "Time", "Defects", "Shape"])
                    plates = self.mock_json_data["leaf_plate_reports"][
                        list(self.mock_json_data["leaf_plate_reports"].keys())[0]
                    ]["plates"]
                    for plate_id, plate_data in plates.items():
                        defects = ", ".join(plate_data["defects"])
                        writer.writerow([
                            plate_id,
                            plate_data["time"],
                            defects,
                            plate_data["shape"]
                        ])
                # Save JSON
                json_filename = self.get_unique_filename(save_directory, ".json")
                with open(json_filename, 'w', encoding='utf-8') as file:
                    json.dump(self.mock_json_data, file, ensure_ascii=False, indent=2)
                messagebox.showinfo(
                    "Success",
                    f"CSV and JSON files saved successfully!\n"
                    f"{os.path.basename(csv_filename)}\n"
                    f"{os.path.basename(json_filename)}"
                )
                dialog.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save files:\n{str(e)}")

    def show_export_dialog(self):
        """Show export options dialog"""
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("Export Options")
        dialog.geometry("280x200")
        dialog.resizable(False, False)
        # Center the dialog
        dialog.transient(self.app)
        dialog.grab_set()
        # Title
        title_label = ctk.CTkLabel(
            dialog,
            text="Choose Export Format",
            font=("Arial", 16, "bold")
        )
        title_label.place(x=140, y=30, anchor="center")

        # CSV Export button: Green
        csv_button = ctk.CTkButton(
            dialog,
            width=200,
            height=35,
            text="Export to CSV",
            font=("Arial", 12, "bold"),
            text_color="#FFFFFF",
            fg_color="#50C878",  # Emerald Green
            hover_color="#27ad60",  # Darker Green
            command=lambda: self.save_to_csv(dialog)
        )
        csv_button.place(x=140, y=70, anchor="center")

        # JSON Export button: Yellow
        json_button = ctk.CTkButton(
            dialog,
            width=200,
            height=35,
            text="Export to JSON",
            font=("Arial", 12, "bold"),
            text_color="#FFFFFF",
            fg_color="#f1c40f",  # Yellow
            hover_color="#d4ac0d",  # Darker Yellow
            command=lambda: self.save_to_json(dialog)
        )
        json_button.place(x=140, y=110, anchor="center")

        # Export Both button: Medium Blue
        both_button = ctk.CTkButton(
            dialog,
            width=200,
            height=35,
            text="Export Both",
            font=("Arial", 12, "bold"),
            text_color="#FFFFFF",
            fg_color="#3498db",  # Medium Blue
            hover_color="#2980b9",  # Darker Blue
            command=lambda: self.save_collected_to_csv_and_json(dialog)
        )
        both_button.place(x=140, y=150, anchor="center")

    # Camera methods
    def start_camera(self):
        """Start camera feed"""
        if self.cap:
            self.camera_running = True
            self.update_camera()

    def stop_camera(self):
        """Stop camera feed"""
        self.camera_running = False
        if self.cap:
            self.cap.release()
            self.cap = None

    def update_camera(self):
        """Update camera feed"""
        if not self.camera_running or not self.cap:
            return
        ret, frame = self.cap.read()
        if not ret:
            self.app.after(30, self.update_camera)
            return
        try:
            frame_resized = cv2.resize(frame, (640, 400))
            # Add detection box
            cv2.rectangle(frame_resized, (20, 20), (620, 380), (0, 255, 0), 2)
            # Convert and display
            frame_rgb = cv2.cvtColor(frame_resized, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(frame_rgb)
            imgtk = ImageTk.PhotoImage(pil_img)
            self.camera_label.configure(image=imgtk, text="")
            self.camera_label.image = imgtk
        except Exception as e:
            print(f"Camera error: {e}")
        self.app.after(30, self.update_camera)

    # Event handlers
    def toggle_data_collection(self):
        """Toggle data collection mode"""
        self.is_collecting_data = not self.is_collecting_data
        if self.is_collecting_data:
            self.toggle_button.configure(
                text="หยุด",
                fg_color="#e74c3c",  # Red for "Stop"
                hover_color="#c0392b"
            )
        else:
            self.toggle_button.configure(
                text="เริ่ม",
                fg_color="#3498db",  # Medium Blue for "Start"
                hover_color="#2980b9"
            )

    def defect_action(self, defect_index):
        """Handle defect action button clicks"""
        defect_name = self.defect_data[defect_index][0]
        messagebox.showinfo("Defect Action", f"Action for: {defect_name}")

    def update_header_time(self):
        """Update header time display"""
        current_time = datetime.now().strftime("%H:%M:%S")
        self.header_time_label.configure(text=current_time)
        self.app.after(1000, self.update_header_time)

    def create_mock_data(self):
        """Create mock data for testing"""
        self.mock_data = {
            "session_start": datetime.now().strftime("%H:%M:%S"),
            "plates_processed": 81,
            "defects_found": 19
        }

    def on_closing(self):
        """Handle application closing"""
        self.stop_camera()
        self.app.destroy()

    def run(self):
        """Start the application"""
        self.app.mainloop()


if __name__ == "__main__":
    app = LeafPlateDetectionApp()
    app.run()