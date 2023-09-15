from fastapi import FastAPI
from main import appends
from fastapi import FastAPI, Form, Request
from fastapi.templating import Jinja2Templates
from package import pack
import models
from database import engine
from fastapi import FastAPI, Form, Request, Depends
import pyautogui
import time

app = FastAPI(openapi_url="/api/openapi.json", docs_url="/api/docs")

templates = Jinja2Templates(directory="templates")
templates.env.globals.update(enumerate=enumerate)

app.include_router(appends, prefix="/api")
app.include_router(pack, prefix="/package")
models.Base.metadata.create_all(bind=engine)

@app.get("/pack/")
def 누른_작동버튼():
    import pyautogui
    import time

    # 작동 버튼의 화면 좌표를 확인하고 클릭할 위치를 지정합니다.
    # 좌표는 웹 브라우저 크기와 위치에 따라 조정해야 합니다.
    # 아래 예제는 스크린 좌표를 사용하므로 웹 브라우저의 위치와 크기에 따라 조정이 필요합니다.
    button_x = 100  # 버튼의 x 좌표
    button_y = 200  # 버튼의 y 좌표

    # 웹 페이지를 열고 작동 버튼을 클릭합니다.
    pyautogui.click(button_x, button_y)

    # 작동 버튼을 클릭한 후 잠시 대기합니다.
    time.sleep(5)  # 예: 2초 동안 대기
    pyautogui.typewrite("Hello, World!", interval=0.1)  # 0.1초 간격으로 문자 입력


@app.get("/")
async def render_upload_form(request: Request):
    return templates.TemplateResponse("main.html", {"request": request})
