from fastapi import FastAPI, Depends, Form, HTTPException
from fastapi.security import OAuth2PasswordRequestForm
from datetime import timedelta

from auth import create_access_token, fake_user
from predict import router as predict_router

app = FastAPI()
app.include_router(predict_router)

@app.post("/token")
def login(form_data: OAuth2PasswordRequestForm = Depends()):
    if form_data.username != fake_user["username"] or form_data.password != fake_user["password"]:
        raise HTTPException(status_code=400, detail="Incorrect username or password")
    
    token = create_access_token(data={"sub": form_data.username}, expires_delta=timedelta(minutes=30))
    return {"access_token": token, "token_type": "bearer"}

@app.get("/")
def read_root():
    return {"msg": "Secure AI API is running"}
