# fine_tune_defects.py
from ultralytics import YOLO
from pathlib import Path
import multiprocessing as mp
import torch, shutil, yaml, glob

# ✅ ใช้ชุด defect และน้ำหนักจากสเตจ shape ตามโฟลเดอร์ที่คุณแสดง
DATA_YAML = Path("dataset2_only_defect/data.yaml")   # ชุดข้อมูลตำหนิ
WEIGHTS   = Path("models/shape_best.pt")             # น้ำหนักจากการเทรนรูปร่าง (stage-1)
OUT_NAME  = "defect_only_from_shape"
EPOCHS    = 100
IMGSZ     = 896
BATCH     = 8
FREEZE    = 10   # แช่แข็ง backbone บางส่วน

def check_labels(data_yaml: Path):
    with open(data_yaml, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    nc = int(cfg.get("nc", 0))
    assert nc == 2, f"data.yaml nc ต้องเป็น 2 (ตอนนี้ = {nc})"
    def _labels_dir(split_key):
        p = cfg[split_key]
        p = (data_yaml.parent / p).resolve()
        return p.parent / "labels"

    bad = []
    for split in ("train", "val", "test"):
        if split not in cfg:
            continue
        labels_dir = _labels_dir(split)
        for txt in labels_dir.glob("**/*.txt"):
            with open(txt, "r", encoding="utf-8") as f:
                for ln, line in enumerate(f, 1):
                    line=line.strip()
                    if not line: 
                        continue
                    try:
                        cid = int(line.split()[0])
                    except Exception:
                        bad.append((txt, ln, line)); continue
                    if cid < 0 or cid >= nc:
                        bad.append((txt, ln, line))
    if bad:
        print("\n[ERROR] พบ class id เกินช่วง 0..1 ในไฟล์ label ต่อไปนี้:")
        for p, ln, line in bad[:30]:
            print(f"  {p} : line {ln} -> {line}")
        raise SystemExit("กรุณาแก้ label ให้มีเฉพาะ 0 (crack), 1 (hole) แล้วรันใหม่ครับ")
    else:
        print("[CHECK] labels OK (class id มีเฉพาะ 0/1)")

def latest_run_dir(pattern="runs/*"):
    cand = sorted(glob.glob(pattern), key=lambda p: Path(p).stat().st_mtime)
    return Path(cand[-1]) if cand else None

def main():
    device = "cuda" if torch.cuda.is_available() else "cpu"
    print(f"[INFO] Using device: {device}")
    assert DATA_YAML.exists(), f"ไม่พบ {DATA_YAML}"
    assert WEIGHTS.exists(),   f"ไม่พบไฟล์ weights: {WEIGHTS}"

    check_labels(DATA_YAML)

    model = YOLO(str(WEIGHTS))

    run_name = f"{OUT_NAME}_img{IMGSZ}_e{EPOCHS}"

    model.train(
        data=str(DATA_YAML),
        epochs=EPOCHS,
        imgsz=IMGSZ,
        batch=BATCH,
        lr0=0.001,
        optimizer="auto",
        close_mosaic=5,
        copy_paste=0.2,
        patience=20,
        workers=2,
        freeze=FREEZE,
        project="runs",
        name=run_name,
        pretrained=True,
        device=device,
        verbose=True,
    )

    model.val(
        data=str(DATA_YAML),
        split="test",
        imgsz=IMGSZ,
        conf=0.25,
        iou=0.6,
        device=device,
    )

    save_dir = latest_run_dir(f"runs/{run_name}*") or latest_run_dir("runs/*")
    if save_dir and (save_dir / "weights/best.pt").exists():
        dst = Path("models/defect_best.pt")
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(save_dir / "weights/best.pt", dst)
        print(f"[OK] คัดลอก best.pt -> {dst}")
    else:
        print("[WARN] หา best.pt ไม่เจอ ลองดูที่โฟลเดอร์ runs/ ด้วยตนเองนะครับ")

if __name__ == "__main__":
    mp.freeze_support()
    main()
