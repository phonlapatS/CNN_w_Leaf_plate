import firebase_admin
from firebase_admin import credentials, db

# 🔹 1. เชื่อม Firebase
cred = credentials.Certificate("credentials.json")
firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://leaf-plate-defect-detec-w-cnn-default-rtdb.asia-southeast1.firebasedatabase.app/'
})

# 🔹 2. ข้อมูลที่จะใส่ (เหมือนที่อยู่ในตารางของคุณ)
date = "26/06/2568"
lot_id = "PTP240627_01"
plate_id = "1"
time = "18:53:00"
defects = ["รอยย่นหรือรอยพับ", "จุดไหม้หรือสีคล้ำ"]  # ถ้าไม่มีตำหนิใช้ []

# 🔹 3. เขียนข้อมูลเข้า Firebase Realtime Database
lot_ref = db.reference(f"/leaf_plate_reports/{lot_id}")
lot_ref.child("date").set(date)  # เก็บวันที่

plate_ref = lot_ref.child("plates").child(plate_id)
plate_ref.set({
    "time": time,
    "defects": defects
})

print(f"✅ Plate {plate_id} inserted successfully in lot {lot_id}")
