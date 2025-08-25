# train_with_time.py
from ultralytics import YOLO
from pathlib import Path
from datetime import timedelta
import time

DATA = r"E:\Final_project\dataset\data.yaml"
MODEL = "yolo11s.pt"
PROJECT = "runs"
RUN_NAME = "leaf_full_quality"
IMGSZ = 640
EPOCHS = 200
PATIENCE = 50

model = YOLO(MODEL)
run_dir = Path(PROJECT) / "detect" / RUN_NAME
run_dir.mkdir(parents=True, exist_ok=True)

t0 = time.perf_counter()

model.train(
    data=DATA, imgsz=IMGSZ, epochs=EPOCHS,
    batch=-1, device=0, patience=PATIENCE,
    project=PROJECT, name=RUN_NAME, save_period=5
)

t1 = time.perf_counter()
elapsed = timedelta(seconds=int(t1 - t0))
summary = f"TOTAL wall-clock: {elapsed}\nRun dir: {run_dir.resolve()}"

print(summary)
(run_dir / "_runtime.txt").write_text(summary, encoding="utf-8")