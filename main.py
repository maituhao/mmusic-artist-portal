import os
import json
import uuid
import shutil
import datetime
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional

import firebase_admin
from firebase_admin import credentials, firestore
import drive_service

# --- Paths ---
BASE_DIR = os.path.dirname(__file__)
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

# --- Firebase Init ---
firebase_creds_str = os.getenv("FIREBASE_CREDS_JSON")
try:
    if firebase_creds_str:
        cred_dict = json.loads(firebase_creds_str)
        cred = credentials.Certificate(cred_dict)
    else:
        cred = credentials.Certificate(os.path.join(BASE_DIR, "firebase_creds.json"))
    firebase_admin.initialize_app(cred)
    db = firestore.client()
except Exception as e:
    print("Warning: Firebase not initialized:", e)
    db = None

# --- Models ---
class TrackInfo(BaseModel):
    title: str
    version: str = "Original"
    mainArtist: str = ""
    featArtist: str = ""
    profile: str = ""
    composer: str = ""
    writer: str = ""
    producer: str = ""
    genre: str = ""
    lang: str = "Vietnamese"
    year: str = ""
    explicit: str = "No"
    tiktok: str = ""
    lyrics: str = ""

class LabelFormRequest(BaseModel):
    productTitle: str
    productType: str
    productArtist: str
    genre: str
    releaseDate: str
    releaseTime: str = ""
    preorderDate: str = ""
    preorderTime: str = ""
    tracks: List[TrackInfo]

# --- Helper functions ---
def get_release(release_id: str):
    if not db:
        return None
    doc = db.collection('releases').document(release_id).get()
    if doc.exists:
        return doc.to_dict()
    return None

def update_release_status(release_id: str, section_key: str):
    if not db:
        return
    doc_ref = db.collection('releases').document(release_id)
    doc = doc_ref.get()
    if doc.exists:
        data = doc.to_dict()
        if "status" not in data:
            data["status"] = {}
        if section_key not in data["status"]:
            data["status"][section_key] = {}
        data["status"][section_key]["completed"] = True
        data["status"][section_key]["timestamp"] = datetime.datetime.now().isoformat()
        doc_ref.set(data)

# --- Routes ---
@app.get("/portal")
async def serve_portal():
    return FileResponse(os.path.join(BASE_DIR, "static", "portal.html"))

@app.get("/api/release/{release_id}")
async def get_release_status(release_id: str):
    data = get_release(release_id)
    if data:
        return JSONResponse({"status": "success", "data": data})
    return JSONResponse({"error": "Không tìm thấy release"}, status_code=404)

@app.post("/api/release/{release_id}/label-form")
async def fill_label_form(release_id: str, payload: LabelFormRequest):
    release_data = get_release(release_id)
    if not release_data:
        return JSONResponse({"error": "Release not found"}, status_code=404)
        
    folder_ids = release_data.get("folder_ids", {})
    target_folder_id = folder_ids.get("1. Label form (track info)")
    
    if not target_folder_id:
        return JSONResponse({"error": "Label form folder not found"}, status_code=500)
    
    template_path = os.path.join(TEMPLATE_DIR, "[MMusic Records] Label Form - Product Info.xlsx")
    if not os.path.exists(template_path):
        return JSONResponse({"error": "Template file not found on server"}, status_code=500)
        
    try:
        import openpyxl
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        
        ws['C4'] = payload.productTitle
        ws['C5'] = payload.productType
        ws['C6'] = payload.productArtist
        ws['C7'] = payload.genre
        ws['C8'] = payload.releaseDate
        if payload.releaseTime:
            ws['C9'] = payload.releaseTime
            
        preorder_value = ""
        if payload.preorderDate:
            preorder_value = payload.preorderDate
            if payload.preorderTime:
                preorder_value += " " + payload.preorderTime
        elif payload.preorderTime:
            preorder_value = payload.preorderTime
        ws['C10'] = preorder_value
        
        start_row = 16
        for i, track in enumerate(payload.tracks):
            row = start_row + i
            ws.cell(row=row, column=1, value=i+1)
            ws.cell(row=row, column=2, value=track.title)
            ws.cell(row=row, column=4, value=track.version)
            ws.cell(row=row, column=5, value=track.mainArtist)
            ws.cell(row=row, column=6, value=track.featArtist)
            ws.cell(row=row, column=8, value=track.profile)
            ws.cell(row=row, column=9, value=track.composer)
            ws.cell(row=row, column=10, value=track.writer)
            ws.cell(row=row, column=11, value=track.producer)
            ws.cell(row=row, column=12, value=track.genre if track.genre else payload.genre)
            ws.cell(row=row, column=13, value=track.lang)
            ws.cell(row=row, column=14, value=track.year)
            ws.cell(row=row, column=15, value="exclusively licensed to MMusic Records")
            ws.cell(row=row, column=16, value="exclusively licensed to MMusic Records")
            ws.cell(row=row, column=19, value=track.explicit)
            ws.cell(row=row, column=20, value=track.tiktok)
            ws.cell(row=row, column=21, value=track.lyrics)
        
        temp_filename = f"/tmp/{release_id}_label_form.xlsx"
        wb.save(temp_filename)
        
        creds = drive_service.get_credentials()
        service = drive_service.get_drive_service(creds)
        
        query = f"'{target_folder_id}' in parents and trashed=false"
        results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])
        
        for item in items:
            try:
                service.files().delete(fileId=item['id']).execute()
            except:
                pass
                
        new_name = f"[MMusic Records] Label Form - {payload.productTitle}.xlsx"
        uploaded = drive_service.upload_file(service, temp_filename, new_name, target_folder_id)
        
        os.remove(temp_filename)
        update_release_status(release_id, "label_form")

        # Lưu metadata label form vào Firestore để Admin trích xuất email
        if db:
            doc_ref = db.collection('releases').document(release_id)
            doc_ref.update({
                "label_form_data": {
                    "productTitle": payload.productTitle,
                    "productType": payload.productType,
                    "productArtist": payload.productArtist,
                    "genre": payload.genre,
                    "releaseDate": payload.releaseDate,
                    "releaseTime": payload.releaseTime,
                    "explicit": payload.tracks[0].explicit if payload.tracks else "No",
                },
                "label_form_link": uploaded.get("webViewLink", "")
            })

        return JSONResponse({"message": "Success", "link": uploaded.get("webViewLink")})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse({"error": str(e)}, status_code=500)

@app.post("/api/release/{release_id}/submit-contract")
async def submit_contract(release_id: str, request: Request):
    release_data = get_release(release_id)
    if not release_data:
        return JSONResponse({"error": "Release not found"}, status_code=404)
        
    folder_ids = release_data.get("folder_ids", {})
    target_folder_id = folder_ids.get("5. Contract Info")
    
    if not target_folder_id:
        return JSONResponse({"error": "Contract Info folder not found. Please re-run automation."}, status_code=500)
    
    try:
        from docx import Document
        from docx.shared import Inches
        
        form = await request.form()
        
        doc = Document()
        doc.add_heading("Hợp Đồng Phát Hành Nhạc", 0)
        
        contract_type = form.get('contractType')
        artist_name = "Unknown_Artist"
        
        if contract_type == "personal":
            doc.add_heading("I. Thông tin khách hàng (cá nhân)", level=1)
            artist_name = form.get('p_name', 'Unknown')
            
            info_lines = [
                f"Họ tên người ký Hợp đồng: {artist_name}",
                f"Nghệ danh: {form.get('p_stagename', '')}",
                f"Sinh ngày: {form.get('p_dob', '')}",
                f"CCCD số: {form.get('p_cccd', '')}",
                f"Cấp ngày: {form.get('p_cccd_date', '')}",
                f"Tại: {form.get('p_cccd_place', '')}",
                f"Địa chỉ hiện tại (cũ): {form.get('p_address_old', '')}",
                f"Địa chỉ hiện tại (mới): {form.get('p_address_new', '')}",
                f"Điện thoại: {form.get('p_phone', '')}",
                f"Email: {form.get('p_email', '')}",
                f"Mã số thuế cá nhân: {form.get('p_tax', '')}",
                f"Số tài khoản: {form.get('p_bank_num', '')}",
                f"Tên tài khoản: {form.get('p_bank_name', '')}",
                f"Ngân hàng: {form.get('p_bank_brand', '')}",
                f"Chi nhánh: {form.get('p_bank_branch', '')}"
            ]
            for line in info_lines:
                doc.add_paragraph(line)
                
            doc.add_heading("Ảnh CCCD:", level=2)
            cccd_front = form.get('p_cccd_front')
            cccd_back = form.get('p_cccd_back')
            
            # Helper func
            async def add_image(file_obj, label):
                if file_obj and hasattr(file_obj, 'filename') and file_obj.filename:
                    temp_img_path = f"/tmp/{uuid.uuid4()}_{file_obj.filename}"
                    with open(temp_img_path, "wb") as buffer:
                        shutil.copyfileobj(file_obj.file, buffer)
                    doc.add_paragraph(f"{label}:")
                    try: doc.add_picture(temp_img_path, width=Inches(4.0))
                    except: doc.add_paragraph("[Không thể chèn ảnh]")
                    os.remove(temp_img_path)
            
            await add_image(cccd_front, "Mặt trước")
            await add_image(cccd_back, "Mặt sau")
            
        elif contract_type == "company":
            doc.add_heading("I. Thông tin khách hàng (công ty)", level=1)
            artist_name = form.get('c_name', 'Unknown_Company')
            
            info_lines = [
                f"Tên Công ty: {artist_name}",
                f"Mã số doanh nghiệp: {form.get('c_tax', '')}",
                f"Địa chỉ trụ sở chính: {form.get('c_address', '')}",
                f"Người đại diện: {form.get('c_rep_name', '')}",
                f"Chức vụ: {form.get('c_rep_title', '')}",
                f"Điện thoại: {form.get('c_phone', '')}",
                f"Email: {form.get('c_email', '')}",
                f"Số tài khoản: {form.get('c_bank_num', '')}",
                f"Tên tài khoản: {form.get('c_bank_name', '')}",
                f"Ngân hàng: {form.get('c_bank_brand', '')}",
                f"Chi nhánh: {form.get('c_bank_branch', '')}"
            ]
            for line in info_lines:
                doc.add_paragraph(line)
        
        temp_doc_path = f"/tmp/Contract_Info_{release_id}.docx"
        doc.save(temp_doc_path)
        
        creds = drive_service.get_credentials()
        service = drive_service.get_drive_service(creds)
        
        query = f"'{target_folder_id}' in parents and name contains 'Thông tin HĐ' and trashed=false"
        results = service.files().list(q=query, spaces='drive', fields='files(id)').execute()
        for item in results.get('files', []):
            try: service.files().delete(fileId=item['id']).execute()
            except: pass
            
        new_name = f"Thông tin HĐ - {artist_name}.docx"
        uploaded = drive_service.upload_file(service, temp_doc_path, new_name, target_folder_id)
        
        os.remove(temp_doc_path)
        update_release_status(release_id, "contract")
        
        return JSONResponse({"message": "Success", "link": uploaded.get("webViewLink")})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse({"error": str(e)}, status_code=500)
