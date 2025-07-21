from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import io
import os
import uuid

app = FastAPI()

# ✅ Enable CORS for frontend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Use ["http://localhost:3000"] in prod
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post('/upload-image/')
async def upload_file(file: UploadFile = File(...)):
    contents = await file.read()

    try:
        img = Image.open(io.BytesIO(contents))
        img.load()
        img = img.convert('RGB')
    except Exception as e:
        return {"error": f"Failed to open image: {str(e)}"}

    # ✅ Resize the image (scale down to reduce Excel size)
    scale_factor = 0.25
    img = img.resize(
        (int(img.width * scale_factor), int(img.height * scale_factor)),
        Image.NEAREST
    )

    pixels = img.load()
    width, height = img.size

    # ✅ Create Excel and color it pixel-by-pixel
    wb = Workbook()
    ws = wb.active

    for y in range(height):
        for x in range(width):
            r, g, b = pixels[x, y]
            hex_color = f'{r:02X}{g:02X}{b:02X}'
            cell = ws.cell(row=y + 1, column=x + 1)
            cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 2
    for row in range(1, height + 1):
        ws.row_dimensions[row].height = 15

    # ✅ Save and return file
    output_filename = f"pixel_art_{uuid.uuid4().hex}.xlsx"
    output_path = os.path.join("/tmp", output_filename)
    wb.save(output_path)

    return FileResponse(
        path=output_path,
        filename="pixel_art.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
