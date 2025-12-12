"""
Simple Image Click - PyAutoGUIで画像を順番にクリックするツール
FastAPI + PyAutoGUI
"""

import os
import time
import json
from pathlib import Path
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel
import shutil
import pyautogui
import pyperclip

app = FastAPI(title="Simple Image Click")

# 設定
IMAGES_DIR = Path(__file__).parent / "images"
TEXTS_FILE = Path(__file__).parent / "texts.json"
DEFAULT_CLICK_INTERVAL = 2.0  # デフォルトのクリック間隔（秒）
DEFAULT_WAIT_TIMEOUT = 30.0  # デフォルトの待機タイムアウト（秒）

# PyAutoGUI設定
pyautogui.FAILSAFE = True  # 画面左上にマウスを移動すると停止


class ActionItem(BaseModel):
    """アクション項目"""
    type: str  # "click", "paste", "wait"
    image_name: str | None = None  # click, wait で使用
    text_index: int | None = None  # paste で使用


class ExecuteRequest(BaseModel):
    """実行リクエスト"""
    actions: list[ActionItem]
    interval: float = DEFAULT_CLICK_INTERVAL
    confidence: float = 0.8
    wait_timeout: float = DEFAULT_WAIT_TIMEOUT


class ExecuteResult(BaseModel):
    """実行結果"""
    success: bool
    message: str
    details: list[dict]


# テキスト管理
def load_texts() -> list[str]:
    """テキスト一覧を読み込む"""
    if not TEXTS_FILE.exists():
        return []
    with open(TEXTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_texts(texts: list[str]):
    """テキスト一覧を保存"""
    with open(TEXTS_FILE, "w", encoding="utf-8") as f:
        json.dump(texts, f, ensure_ascii=False, indent=2)


@app.get("/", response_class=HTMLResponse)
async def root():
    """メインページを返す"""
    html_path = Path(__file__).parent / "index.html"
    return FileResponse(html_path)


@app.get("/api/images")
async def get_images():
    """imagesフォルダ内の画像一覧を返す"""
    if not IMAGES_DIR.exists():
        IMAGES_DIR.mkdir(parents=True, exist_ok=True)
        return {"images": []}

    image_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}
    images = []
    for file in IMAGES_DIR.iterdir():
        if file.suffix.lower() in image_extensions:
            images.append({
                "name": file.name,
                "path": f"/images/{file.name}"
            })

    return {"images": sorted(images, key=lambda x: x["name"])}


@app.get("/api/texts")
async def get_texts():
    """テキスト一覧を返す"""
    return {"texts": load_texts()}


@app.post("/api/texts")
async def add_text(data: dict):
    """テキストを追加"""
    text = data.get("text", "").strip()
    if not text:
        raise HTTPException(status_code=400, detail="テキストが空です")

    texts = load_texts()
    texts.append(text)
    save_texts(texts)
    return {"success": True, "index": len(texts) - 1, "texts": texts}


@app.put("/api/texts/{index}")
async def update_text(index: int, data: dict):
    """テキストを更新"""
    texts = load_texts()
    if index < 0 or index >= len(texts):
        raise HTTPException(status_code=404, detail="テキストが見つかりません")

    text = data.get("text", "").strip()
    if not text:
        raise HTTPException(status_code=400, detail="テキストが空です")

    texts[index] = text
    save_texts(texts)
    return {"success": True, "texts": texts}


@app.delete("/api/texts/{index}")
async def delete_text(index: int):
    """テキストを削除"""
    texts = load_texts()
    if index < 0 or index >= len(texts):
        raise HTTPException(status_code=404, detail="テキストが見つかりません")

    texts.pop(index)
    save_texts(texts)
    return {"success": True, "texts": texts}


@app.get("/api/settings")
async def get_settings():
    """現在の設定を返す"""
    return {
        "default_interval": DEFAULT_CLICK_INTERVAL,
        "confidence": 0.8,
        "wait_timeout": DEFAULT_WAIT_TIMEOUT
    }


@app.post("/api/upload")
async def upload_image(file: UploadFile = File(...)):
    """画像をアップロードする"""
    image_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}
    ext = Path(file.filename).suffix.lower()
    if ext not in image_extensions:
        raise HTTPException(status_code=400, detail=f"対応していないファイル形式です: {ext}")

    IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    file_path = IMAGES_DIR / file.filename

    if file_path.exists():
        base = Path(file.filename).stem
        counter = 1
        while file_path.exists():
            file_path = IMAGES_DIR / f"{base}_{counter}{ext}"
            counter += 1

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {
        "success": True,
        "filename": file_path.name,
        "message": f"アップロード成功: {file_path.name}"
    }


@app.delete("/api/images/{image_name}")
async def delete_image(image_name: str):
    """画像を削除する"""
    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        raise HTTPException(status_code=404, detail=f"画像が見つかりません: {image_name}")

    image_path.unlink()
    return {"success": True, "message": f"削除しました: {image_name}"}


@app.post("/api/execute", response_model=ExecuteResult)
async def execute_actions(request: ExecuteRequest):
    """アクションを順番に実行する"""
    if not request.actions:
        raise HTTPException(status_code=400, detail="アクションが選択されていません")

    texts = load_texts()
    results = []
    success_count = 0

    for i, action in enumerate(request.actions):
        try:
            if action.type == "click":
                # クリック
                result = execute_click(action.image_name, request.confidence)
            elif action.type == "paste":
                # テキスト貼り付け
                result = execute_paste(action.text_index, texts)
            elif action.type == "wait":
                # 画像が出るまで待機
                result = execute_wait(action.image_name, request.confidence, request.wait_timeout)
            else:
                result = {"status": "error", "message": f"不明なアクション: {action.type}"}

            results.append(result)
            if result["status"] == "success":
                success_count += 1

            # 次のアクションまで待機（最後以外）
            if i < len(request.actions) - 1 and result["status"] == "success":
                time.sleep(request.interval)

        except Exception as e:
            results.append({"status": "error", "message": f"エラー: {str(e)}"})

    return ExecuteResult(
        success=success_count == len(request.actions),
        message=f"{success_count}/{len(request.actions)} 件のアクションが成功しました",
        details=results
    )


def execute_click(image_name: str, confidence: float) -> dict:
    """画像をクリック"""
    if not image_name:
        return {"status": "error", "message": "画像が指定されていません"}

    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        return {"status": "error", "message": f"画像ファイルが見つかりません: {image_name}"}

    location = pyautogui.locateCenterOnScreen(str(image_path), confidence=confidence)
    if location is None:
        return {"status": "not_found", "message": f"画面上に画像が見つかりません: {image_name}"}

    pyautogui.click(location)
    return {"status": "success", "message": f"[クリック] {image_name} (位置: {location})"}


def execute_paste(text_index: int, texts: list[str]) -> dict:
    """テキストを貼り付け"""
    if text_index is None or text_index < 0 or text_index >= len(texts):
        return {"status": "error", "message": f"テキストが見つかりません: index={text_index}"}

    text = texts[text_index]
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')
    return {"status": "success", "message": f"[貼付] [{text_index + 1}] {text[:30]}..."}


def execute_wait(image_name: str, confidence: float, timeout: float) -> dict:
    """画像が出るまで待機"""
    if not image_name:
        return {"status": "error", "message": "画像が指定されていません"}

    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        return {"status": "error", "message": f"画像ファイルが見つかりません: {image_name}"}

    start_time = time.time()
    while time.time() - start_time < timeout:
        location = pyautogui.locateCenterOnScreen(str(image_path), confidence=confidence)
        if location is not None:
            return {"status": "success", "message": f"[待機] 画像を検出: {image_name} (位置: {location})"}
        time.sleep(0.5)

    return {"status": "timeout", "message": f"タイムアウト: {image_name} が {timeout}秒以内に見つかりませんでした"}


# 後方互換性のため古いAPIも残す
class ClickRequest(BaseModel):
    image_names: list[str]
    interval: float = DEFAULT_CLICK_INTERVAL
    confidence: float = 0.8


@app.post("/api/click")
async def execute_clicks(request: ClickRequest):
    """画像を順番にクリックする（後方互換）"""
    actions = [ActionItem(type="click", image_name=name) for name in request.image_names]
    exec_request = ExecuteRequest(
        actions=actions,
        interval=request.interval,
        confidence=request.confidence
    )
    return await execute_actions(exec_request)


# 画像ファイルを静的ファイルとして配信
app.mount("/images", StaticFiles(directory=str(IMAGES_DIR)), name="images")


if __name__ == "__main__":
    import uvicorn
    print("=" * 50)
    print("Simple Image Click Server")
    print("=" * 50)
    print(f"画像フォルダ: {IMAGES_DIR}")
    print("ブラウザで http://localhost:8000 を開いてください")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8000)
