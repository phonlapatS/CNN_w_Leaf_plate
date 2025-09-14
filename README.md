<img width="1909" height="1080" alt="image" src="https://github.com/user-attachments/assets/a18476d7-74b5-4632-a59e-6b03796123b5" />


# How to setup this project
** this project use python version 3.10.7 **
### Build Environment for python
python -m venv yolo-env
name....\Scripts\activate

### upgrade pip before install lib
pip install --upgrade pip
### install ultralytics for use yolov11
pip install ultralytics==8.*
#### เช็ค yolo version
yolo --version

pip install opencv-python
pip install numpy

### ติดตั้ง lib นี้เพื่อเปิดใช้ CUDA หรือ เอา GPU มาช่วยเทรน
pip install --index-url https://download.pytorch.org/whl/cu121 --no-cache-dir torch
pip install --index-url https://download.pytorch.org/whl/cu121 --no-cache-dir torchvision
pip install --index-url https://download.pytorch.org/whl/cu121 --no-cache-dir torchaudio

หรือ

pip install --index-url https://download.pytorch.org/whl/cu121 --no-cache-dir torch torchvision torchaudio


### ทดสอบว่าเจอ GPU รึยัง
python - << 'PY'
import torch
print("Torch:", torch.__version__)
print("CUDA available:", torch.cuda.is_available())
if torch.cuda.is_available():
    print("GPU:", torch.cuda.get_device_name(0))
PY

# lib อื่นๆ
### ติดตั้งรวดเดียว
pip install --upgrade customtkinter opencv-python Pillow numpy

### ติดตั้งทีละอัน
pip install --upgrade customtkinter
pip install --upgrade opencv-python
pip install --upgrade Pillow
pip install --upgrade numpy
