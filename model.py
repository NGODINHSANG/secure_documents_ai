import torch
import torchvision.transforms as transforms
from torchvision import models

from PIL import Image

device = torch.device("cuda" if torch.cuda.is_available() else "cpu")

# Load pretrained model
model = models.resnet18(pretrained=True)
model.eval()
model.to(device)

transform = transforms.Compose([
    transforms.Resize((224, 224)),
    transforms.ToTensor()
])

def predict_image(image: Image.Image) -> str:
    img = transform(image).unsqueeze(0).to(device)
    with torch.no_grad():
        output = model(img)
        pred = output.argmax(1).item()
    return str(pred)
