from fastapi import APIRouter, UploadFile, File, Depends, HTTPException
from PIL import Image
from io import BytesIO

from model import predict_image
from auth import verify_token

router = APIRouter()

@router.post("/predict")
async def predict(file: UploadFile = File(...), user: str = Depends(verify_token)):
    if file.content_type not in ["image/jpeg", "image/png"]:
        raise HTTPException(status_code=400, detail="Invalid image format")
    
    image_data = await file.read()
    image = Image.open(BytesIO(image_data)).convert("RGB")
    label = predict_image(image)

    return {"user": user, "label": label}
