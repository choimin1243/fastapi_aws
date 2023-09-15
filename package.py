from fastapi.templating import Jinja2Templates
from fastapi import FastAPI, Form, Request, Depends
from fastapi import APIRouter,Request
from sqlalchemy.orm import Session
from database import SessionLocal
from models import Todos


pack = APIRouter()
templates = Jinja2Templates(directory="templates")
templates.env.globals.update(enumerate=enumerate)

def get_db():
    db=SessionLocal()
    try:
        yield db
    finally:
        db.close()



@pack.get("/")
async def render_upload_form(request: Request):
    return templates.TemplateResponse("package.html", {"request": request})


def hello(db: Session = Depends(get_db)):
    packages=get_all_todos_from_db(db)
    package=[package.title for package in packages]
    return package


@pack.post("/send-text/")
async def send_list(request: Request, db: Session = Depends(get_db), text: str = Form(...)):
    print(text)
    new_todo = Todos(title=text)
    db.add(new_todo)
    db.commit()
    packages=get_all_todos_from_db(db)
    package=[package.title for package in packages]
    print(package)
    return templates.TemplateResponse("package.html", {"request": request})




def get_all_todos_from_db(db: Session):
    todos = db.query(Todos).all()
    return todos
