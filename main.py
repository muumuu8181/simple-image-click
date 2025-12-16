"""
Simple Image Click - PyAutoGUIで画像を順番にクリックするツール
FastAPI + PyAutoGUI
"""

import os
import time
import json
import random
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
FLOWS_FILE = Path(__file__).parent / "flows.json"  # アクションフロー保存
LOG_FILE = Path(__file__).parent / "batch_log.txt"  # バッチ実行ログ
DEFAULT_CLICK_INTERVAL = 2.0  # デフォルトのクリック間隔（秒）
DEFAULT_WAIT_TIMEOUT = 30.0  # デフォルトの待機タイムアウト（秒）

# PyAutoGUI設定
pyautogui.FAILSAFE = True  # 画面左上にマウスを移動すると停止


class ActionItem(BaseModel):
    """アクション項目"""
    type: str  # "click", "paste", "wait", "click_or"
    image_name: str | None = None  # click, wait で使用
    image_names: list[str] | None = None  # click_or で使用（複数画像）
    text_id: str | None = None  # paste で使用（8桁ID）
    text_index: int | None = None  # 後方互換用（非推奨）


class ExecuteRequest(BaseModel):
    """実行リクエスト"""
    actions: list[ActionItem]
    interval: float = DEFAULT_CLICK_INTERVAL
    confidence: float = 0.8
    wait_timeout: float = DEFAULT_WAIT_TIMEOUT
    cursor_speed: float = 0.5  # カーソル移動速度（秒）


class ExecuteResult(BaseModel):
    """実行結果"""
    success: bool
    message: str
    details: list[dict]


# テキスト管理（ID付き形式: {id: {id, text, created_at}}）
def generate_text_id() -> str:
    """8桁のユニークIDを生成"""
    return str(random.randint(10000000, 99999999))


def load_texts() -> dict:
    """テキスト一覧を読み込む（ID付き辞書形式）"""
    if not TEXTS_FILE.exists():
        return {}
    with open(TEXTS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
        # 旧形式（リスト）からの移行対応
        if isinstance(data, list):
            return migrate_texts_to_id_format(data)
        return data


def migrate_texts_to_id_format(old_texts: list[str]) -> dict:
    """旧形式（リスト）から新形式（ID付き辞書）に移行"""
    new_texts = {}
    for text in old_texts:
        text_id = generate_text_id()
        while text_id in new_texts:
            text_id = generate_text_id()
        new_texts[text_id] = {
            "id": text_id,
            "text": text,
            "created_at": time.strftime("%Y-%m-%d %H:%M:%S")
        }
    # 移行後は新形式で保存
    save_texts(new_texts)
    return new_texts


def save_texts(texts: dict):
    """テキスト一覧を保存"""
    with open(TEXTS_FILE, "w", encoding="utf-8") as f:
        json.dump(texts, f, ensure_ascii=False, indent=2)


def get_text_by_id(texts: dict, text_id: str) -> str | None:
    """IDからテキストを取得"""
    if text_id in texts:
        return texts[text_id]["text"]
    return None


# フロー管理
def load_flows() -> dict:
    """フロー一覧を読み込む"""
    if not FLOWS_FILE.exists():
        return {}
    with open(FLOWS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_flows(flows: dict):
    """フロー一覧を保存"""
    with open(FLOWS_FILE, "w", encoding="utf-8") as f:
        json.dump(flows, f, ensure_ascii=False, indent=2)


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
    text_id = generate_text_id()
    while text_id in texts:
        text_id = generate_text_id()

    texts[text_id] = {
        "id": text_id,
        "text": text,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    save_texts(texts)
    return {"success": True, "id": text_id, "texts": texts}


@app.put("/api/texts/{text_id}")
async def update_text(text_id: str, data: dict):
    """テキストを更新"""
    texts = load_texts()
    if text_id not in texts:
        raise HTTPException(status_code=404, detail="テキストが見つかりません")

    text = data.get("text", "").strip()
    if not text:
        raise HTTPException(status_code=400, detail="テキストが空です")

    texts[text_id]["text"] = text
    save_texts(texts)
    return {"success": True, "texts": texts}


@app.delete("/api/texts/{text_id}")
async def delete_text(text_id: str):
    """テキストを削除"""
    texts = load_texts()
    if text_id not in texts:
        raise HTTPException(status_code=404, detail="テキストが見つかりません")

    del texts[text_id]
    save_texts(texts)
    return {"success": True, "texts": texts}


# フローAPI
@app.get("/api/flows")
async def get_flows():
    """フロー一覧を返す"""
    return {"flows": load_flows()}


@app.post("/api/flows")
async def save_flow(data: dict):
    """フローを保存"""
    name = data.get("name", "").strip()
    actions = data.get("actions", [])

    if not name:
        raise HTTPException(status_code=400, detail="フロー名が空です")
    if not actions:
        raise HTTPException(status_code=400, detail="アクションが空です")

    flows = load_flows()
    flows[name] = {
        "actions": actions,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    save_flows(flows)
    return {"success": True, "flows": flows}


@app.delete("/api/flows/{flow_name}")
async def delete_flow(flow_name: str):
    """フローを削除"""
    flows = load_flows()
    if flow_name not in flows:
        raise HTTPException(status_code=404, detail=f"フローが見つかりません: {flow_name}")

    del flows[flow_name]
    save_flows(flows)
    return {"success": True, "flows": flows}


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


@app.put("/api/images/{image_name}")
async def replace_image(image_name: str, file: UploadFile = File(...)):
    """既存の画像を差し替える"""
    file_path = IMAGES_DIR / image_name
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"画像が見つかりません: {image_name}")

    image_extensions = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}
    ext = Path(file.filename).suffix.lower()
    if ext not in image_extensions:
        raise HTTPException(status_code=400, detail=f"対応していないファイル形式です: {ext}")

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {
        "success": True,
        "filename": image_name,
        "message": f"差し替え成功: {image_name}"
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
            elif action.type == "click_or":
                # クリック（OR条件：最初に見つかった画像をクリック）
                result = execute_click_or(action.image_names, request.confidence)
            elif action.type == "paste":
                # テキスト貼り付け（IDベース）
                result = execute_paste(action.text_id, texts)
            elif action.type == "wait":
                # 画像が出るまで待機
                result = execute_wait(action.image_name, request.confidence, request.wait_timeout, request.cursor_speed)
            else:
                result = {"status": "error", "message": f"不明なアクション: {action.type}"}

            results.append(result)
            if result["status"] == "success":
                success_count += 1

            # 次のアクションまで待機（最後以外）
            if i < len(request.actions) - 1 and result["status"] == "success":
                time.sleep(request.interval)

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(f"エラー詳細: {error_detail}")  # サーバーログに出力
            results.append({"status": "error", "message": f"エラー: {type(e).__name__}: {str(e)}"})

    return ExecuteResult(
        success=success_count == len(request.actions),
        message=f"{success_count}/{len(request.actions)} 件のアクションが成功しました",
        details=results
    )


def find_best_match_confidence(image_path: str, target_confidence: float) -> tuple[float | None, any]:
    """画像の最も近いマッチの信頼度を調べる（段階的に閾値を下げて検索）"""
    # 設定した閾値で見つかるかチェック
    try:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=target_confidence)
        if location is not None:
            return target_confidence, location
    except pyautogui.ImageNotFoundException:
        pass
    except Exception:
        pass

    # 見つからない場合、閾値を下げて検索
    test_confidences = [0.7, 0.6, 0.5, 0.4, 0.3]
    for conf in test_confidences:
        if conf >= target_confidence:
            continue
        try:
            location = pyautogui.locateCenterOnScreen(image_path, confidence=conf)
            if location is not None:
                # この閾値で見つかった = 実際の信頼度はこの値以上、次の値未満
                return conf, location
        except pyautogui.ImageNotFoundException:
            pass
        except Exception:
            pass

    return None, None


def execute_click(image_name: str, confidence: float) -> dict:
    """画像をクリック"""
    if not image_name:
        return {"status": "error", "message": "画像が指定されていません"}

    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        return {"status": "error", "message": f"画像ファイルが見つかりません: {image_name}"}

    found_conf, location = find_best_match_confidence(str(image_path), confidence)

    if found_conf is not None and found_conf >= confidence:
        pyautogui.click(location)
        return {"status": "success", "message": f"[クリック] {image_name} (位置: {location})"}

    if found_conf is not None:
        return {"status": "not_found", "message": f"画面上に画像が見つかりません: {image_name} (最大{int(found_conf*100)}%で検出、設定は{int(confidence*100)}%)"}

    return {"status": "not_found", "message": f"画面上に画像が見つかりません: {image_name} (30%未満、画像が画面にない可能性)"}


def execute_click_or(image_names: list[str], confidence: float) -> dict:
    """複数画像のうち最初に見つかった画像をクリック"""
    if not image_names or len(image_names) == 0:
        return {"status": "error", "message": "画像が指定されていません"}

    not_found_details = []
    for image_name in image_names:
        image_path = IMAGES_DIR / image_name
        if not image_path.exists():
            not_found_details.append(f"{image_name}(ファイルなし)")
            continue

        found_conf, location = find_best_match_confidence(str(image_path), confidence)

        if found_conf is not None and found_conf >= confidence:
            pyautogui.click(location)
            return {"status": "success", "message": f"[クリックOR] {image_name} (位置: {location})"}

        if found_conf is not None:
            not_found_details.append(f"{image_name}({int(found_conf*100)}%)")
        else:
            not_found_details.append(f"{image_name}(30%未満)")

    # どの画像も見つからなかった
    detail_str = " / ".join(not_found_details)
    return {"status": "not_found", "message": f"画面上にどの画像も見つかりません: {detail_str} (設定: {int(confidence*100)}%)"}


def execute_paste(text_id: str, texts: dict) -> dict:
    """テキストを貼り付け（IDベース）"""
    if text_id is None or text_id not in texts:
        return {"status": "error", "message": f"テキストが見つかりません: ID={text_id}"}

    text = texts[text_id]["text"]
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')
    return {"status": "success", "message": f"[貼付] [ID:{text_id}] {text[:30]}..."}


def smooth_move_cursor(target_x: int, target_y: int, duration: float = 0.5):
    """カーソルをスムーズに移動"""
    # PyAutoGUIのmoveToにdurationとtweenを指定
    # MINIMUM_DURATION(0.1秒)以上必要
    actual_duration = max(duration, 0.2)
    pyautogui.moveTo(target_x, target_y, duration=actual_duration, tween=pyautogui.easeInOutQuad)


def execute_wait(image_name: str, confidence: float, timeout: float, cursor_speed: float = 0.5) -> dict:
    """画像が出るまで待機（カーソルを小さく動かして待機中を示す）"""
    if not image_name:
        return {"status": "error", "message": "画像が指定されていません"}

    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        return {"status": "error", "message": f"画像ファイルが見つかりません: {image_name}"}

    start_time = time.time()
    move_direction = 1  # カーソル移動方向（1: 右, -1: 左）
    move_amount = 100  # 移動量（ピクセル）- 見やすく

    while time.time() - start_time < timeout:
        location = pyautogui.locateCenterOnScreen(str(image_path), confidence=confidence)
        if location is not None:
            return {"status": "success", "message": f"[待機] 画像を検出: {image_name} (位置: {location})"}

        # 待機中を示すためにカーソルを左右にスムーズに動かす
        current_pos = pyautogui.position()
        target_x = current_pos[0] + (move_amount * move_direction)
        smooth_move_cursor(target_x, current_pos[1], cursor_speed)
        move_direction *= -1  # 方向を反転

        time.sleep(0.1)  # 画像チェックの間隔

    # タイムアウト時に信頼度を調べる
    found_conf, _ = find_best_match_confidence(str(image_path), confidence)
    if found_conf is not None:
        return {"status": "timeout", "message": f"タイムアウト: {image_name} が {timeout}秒以内に見つかりませんでした (最大{int(found_conf*100)}%、設定は{int(confidence*100)}%)"}
    return {"status": "timeout", "message": f"タイムアウト: {image_name} が {timeout}秒以内に見つかりませんでした (30%未満、画像が画面にない可能性)"}


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


# ログ保存API
class LogRequest(BaseModel):
    log: str

@app.post("/api/log")
async def save_log(request: LogRequest):
    """ログをファイルに追記"""
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"\n{'='*60}\n")
        f.write(f"[{timestamp}]\n")
        f.write(request.log)
        f.write("\n")
    return {"success": True}


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
