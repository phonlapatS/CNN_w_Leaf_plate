# train_shape.py
from ultralytics import YOLO
from pathlib import Path
import torch, os

# ---------------------------------------
# CONFIG (แก้ได้ตามชุดของคุณ)
# ---------------------------------------
DATA_YAML    = Path("dataset2_only_shape/data.yaml")    # มีเฉพาะคลาสรูปร่างจาน
WEIGHTS_S    = Path("yolov11s.pt")                      # เริ่มใหม่จาก COCO
WEIGHTS_OLD  = Path("models/shape_best.pt")             # น้ำหนักเดิมของคุณ (ถ้ามี)
RUN_PREFIX   = "shape_only_from_dataset2"

# แกนหลักสำหรับ Stage-A (เฟรมเต็ม)
EPOCHS   = 100                 # 80–120 พอ
IMGSZ    = 896                 # 768–960 ดีสุดเรื่องความไว/ความแม่น → ใช้ 896
BATCH    = 16                  # OOM ให้ลดเป็น 12/8
LR0      = 0.003               # เริ่มสูงนิด + cosine schedule
FREEZE   = 0                   # ให้ทั้งโมเดลเรียนรู้ (จะได้ unlearn ฉากเก่าได้ถ้าฟื้นจาก shape_best)
PATIENCE = 25
WORKERS  = 8

# เลือก device อัตโนมัติ
DEVICE = 0 if torch.cuda.is_available() else "cpu"

# ---------------------------------------
# เลือกจุดเริ่ม: ถ้ามี shape_best.pt จะเริ่มจากของเดิม
# ถ้าอยาก “เริ่มใหม่จาก yolov11s” ให้เปลี่ยน init_path = WEIGHTS_S
# ---------------------------------------
init_path = WEIGHTS_OLD if WEIGHTS_OLD.exists() else WEIGHTS_S
print(f"[INFO] INIT_WEIGHTS = {init_path}")

# โหลดโมเดล
model = YOLO(str(init_path))

# ---------------------------------------
# TRAIN
# ---------------------------------------
model.train(
    data=str(DATA_YAML),
    epochs=EPOCHS,
    imgsz=IMGSZ,
    batch=BATCH,
    lr0=LR0,
    cos_lr=True,              # learning rate แบบ cosine → converge นิ่ม
    patience=PATIENCE,        # early stop
    workers=WORKERS,
    device=DEVICE,
    cache="ram",              # โหลดไวขึ้น
    rect=True,                # pack รูปหลายอัตราส่วนให้ดี
    # ---- augmentation ที่เหมาะกับ "รูปร่างจาน" บนเฟรมเต็ม ----
    mosaic=0.8,               # ช่วงต้นช่วย generalize
    close_mosaic=15,          # ปิดก่อนจบ ~15 epochs เพื่อโฟกัสภาพจริง
    copy_paste=0.0,           # งานรูปทรงไม่จำเป็น
    degrees=5.0,              # เผื่อกล้องเอียงเล็กน้อย
    translate=0.05,           # ขยับ 5%
    scale=0.10,               # ย่อ/ขยายเล็กน้อย
    shear=0.0, perspective=0.0,
    fliplr=0.5, flipud=0.0,   # ซ้าย-ขวาได้ พอ
    hsv_h=0.015, hsv_s=0.5, hsv_v=0.3,  # แกว่งสี/แสงพอประมาณ
    amp=True,                 # mixed precision ให้ไวขึ้น
    pretrained=True,
    freeze=FREEZE,            # 0 = ให้ทั้ง backbone ปรับตัวได้
    # ---- บันทึกผล ----
    project="runs",
    name=f"{RUN_PREFIX}_img{IMGSZ}_e{EPOCHS}",
    exist_ok=True,
    seed=42,
    verbose=True,
)

# ---------------------------------------
# VALIDATE บน test split ของ shape
# ---------------------------------------
model.val(
    data=str(DATA_YAML),
    split="test",
    imgsz=IMGSZ,
    conf=0.25,       # ค่าเริ่มต้นโอเคสำหรับ Stage-A
    iou=0.6,         # ให้กรอบกระชับขึ้นนิด
    device=DEVICE,
    verbose=True,
)
