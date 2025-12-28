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
import pygetwindow as gw

# ブラウザウィンドウ制御用
saved_window_state = None

def minimize_browser_window():
    """ブラウザウィンドウを最小化して状態を保存"""
    global saved_window_state
    try:
        # localhost:8000を含むウィンドウを探す
        windows = gw.getWindowsWithTitle('Simple Image Click')
        if not windows:
            windows = gw.getWindowsWithTitle('localhost:8000')
        if windows:
            win = windows[0]
            saved_window_state = {
                'left': win.left,
                'top': win.top,
                'width': win.width,
                'height': win.height,
                'window': win
            }
            win.minimize()
            print(f'[INFO] ウィンドウを最小化しました')
    except Exception as e:
        print(f'[WARN] ウィンドウ最小化失敗: {e}')

def restore_browser_window():
    """ブラウザウィンドウを復元"""
    global saved_window_state
    try:
        if saved_window_state and saved_window_state.get('window'):
            win = saved_window_state['window']
            win.restore()
            print(f'[INFO] ウィンドウを復元しました')
            saved_window_state = None
    except Exception as e:
        print(f'[WARN] ウィンドウ復元失敗: {e}')

app = FastAPI(title="Simple Image Click")

# 実行状態管理
import threading
import uuid

class ExecutionState:
    def __init__(self):
        self.is_running = False
        self.abort_flag = False
        self.execution_id = None
        self.current_step = 0
        self.total_steps = 0
        self.results = []
        self.completed = False
        self.lock = threading.Lock()

    def start(self, total_steps: int) -> str:
        with self.lock:
            if self.is_running:
                return None  # 既に実行中
            self.is_running = True
            self.abort_flag = False
            self.execution_id = str(uuid.uuid4())[:8]
            self.current_step = 0
            self.total_steps = total_steps
            self.results = []
            self.completed = False
            return self.execution_id

    def add_result(self, result: dict):
        with self.lock:
            self.results.append(result)
            self.current_step = len(self.results)

    def finish(self):
        with self.lock:
            self.completed = True
            self.is_running = False

    def abort(self):
        with self.lock:
            self.abort_flag = True

    def get_status(self) -> dict:
        with self.lock:
            return {
                "is_running": self.is_running,
                "execution_id": self.execution_id,
                "current_step": self.current_step,
                "total_steps": self.total_steps,
                "results": list(self.results),
                "completed": self.completed,
                "aborted": self.abort_flag
            }

execution_state = ExecutionState()

# 後方互換用（古いコードで参照している場合）
execution_abort_flag = False

# 設定
IMAGES_DIR = Path(__file__).parent / "images"
TEXTS_FILE = Path(__file__).parent / "texts.json"
FLOWS_FILE = Path(__file__).parent / "flows.json"  # アクションフロー保存
LOG_FILE = Path(__file__).parent / "batch_log.txt"  # バッチ実行ログ
DEFAULT_CLICK_INTERVAL = 2.0  # デフォルトのクリック間隔（秒）
DEFAULT_WAIT_TIMEOUT = 1800.0  # デフォルトの待機タイムアウト（秒）= 30分

# PyAutoGUI設定
pyautogui.FAILSAFE = True  # 画面左上にマウスを移動すると停止


class ActionItem(BaseModel):
    """アクション項目"""
    type: str  # "click", "paste", "paste_fixed", "wait", "click_or", "wait_disappear", "wait_seconds", "pagedown", "save_to_file"
    image_name: str | None = None  # click, wait, wait_disappear で使用
    image_names: list[str] | None = None  # click_or で使用（複数画像）
    text_id: str | None = None  # paste, save_to_file で使用（8桁ID）
    text: str | None = None  # paste_fixed で使用（固定テキスト）
    text_index: int | None = None  # 後方互換用（非推奨）
    seconds: float | None = None  # wait_seconds で使用
    count: int | None = None  # pagedown で使用（回数）
    flow_name: str | None = None  # save_to_file で使用（フロー名）- ファイル内ヘッダー用
    group_name: str | None = None  # save_to_file で使用（グループ名）- ファイル名用


class ExecuteRequest(BaseModel):
    """実行リクエスト"""
    actions: list[ActionItem]
    interval: float = DEFAULT_CLICK_INTERVAL
    confidence: float = 0.95
    min_confidence: float = 0.7  # 最低認識精度（ここまで下げて試す）
    wait_timeout: float = DEFAULT_WAIT_TIMEOUT
    cursor_speed: float = 0.5  # カーソル移動速度（秒）
    start_delay: float = 0.0  # 開始前待機（秒）- 最初のアクション前に待つ
    flow_name: str | None = None  # フロー名（Grok日本語対応などで使用）


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
    group = data.get("group", "")

    if not name:
        raise HTTPException(status_code=400, detail="フロー名が空です")
    if not actions:
        raise HTTPException(status_code=400, detail="アクションが空です")

    flows = load_flows()
    flows[name] = {
        "actions": actions,
        "group": group,
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

@app.put("/api/flows/{flow_name}/group")
async def change_flow_group(flow_name: str, data: dict):
    """フローのグループを変更"""
    flows = load_flows()
    if flow_name not in flows:
        raise HTTPException(status_code=404, detail=f"フローが見つかりません: {flow_name}")

    new_group = data.get("group", "")
    flows[flow_name]["group"] = new_group

    # save_to_fileアクションのgroup_nameも更新
    for action in flows[flow_name].get("actions", []):
        if action.get("type") == "save_to_file":
            action["group_name"] = new_group

    save_flows(flows)
    return {"success": True, "flows": flows}

@app.put("/api/flows/{flow_name}/suspend")
async def toggle_flow_suspend(flow_name: str, data: dict):
    """フローの休止状態を変更"""
    flows = load_flows()
    if flow_name not in flows:
        raise HTTPException(status_code=404, detail=f"フローが見つかりません: {flow_name}")

    suspended = data.get("suspended", False)
    flows[flow_name]["suspended"] = suspended

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


@app.post("/api/abort")
async def abort_execution():
    """実行を中止"""
    global execution_abort_flag
    execution_abort_flag = True
    execution_state.abort()
    return {"success": True, "message": "中止フラグを設定しました"}


@app.get("/api/execute/status")
async def get_execution_status():
    """実行状態を取得"""
    return execution_state.get_status()


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


def run_actions_in_background(request_dict: dict):
    """バックグラウンドでアクションを実行"""
    global execution_abort_flag

    texts = load_texts()
    actions = request_dict["actions"]
    confidence = request_dict.get("confidence", 0.95)
    min_confidence = request_dict.get("min_confidence", 0.7)
    wait_timeout = request_dict.get("wait_timeout", DEFAULT_WAIT_TIMEOUT)
    cursor_speed = request_dict.get("cursor_speed", 0.5)
    interval = request_dict.get("interval", DEFAULT_CLICK_INTERVAL)
    start_delay = request_dict.get("start_delay", 0.0)

    # フロー名を取得（リクエストから直接、またはsave_to_fileアクションから）
    flow_name_for_paste = request_dict.get('flow_name')
    print(f"[DEBUG] execute: flow_name from request = {flow_name_for_paste}")
    if not flow_name_for_paste:
        # 後方互換: save_to_fileアクションからflow_nameを取得
        for a in actions:
            if a.get('type') == 'save_to_file' and a.get('flow_name'):
                flow_name_for_paste = a.get('flow_name')
                break

    # ブラウザウィンドウを最小化
    minimize_browser_window()
    
    # 開始前待機
    if start_delay > 0:
        execution_state.add_result({"status": "info", "message": f"[開始待機] {start_delay}秒待機中..."})
        time.sleep(start_delay)
        # 待機後に中止されていないか確認
        if execution_state.abort_flag or execution_abort_flag:
            execution_state.add_result({"status": "aborted", "message": f"[開始待機] 中止されました"})
            restore_browser_window()
            execution_state.finish()
            return

    for i, action_dict in enumerate(actions):
        # 中止チェック
        if execution_state.abort_flag or execution_abort_flag:
            execution_state.add_result({"status": "aborted", "message": f"[中止] ユーザーにより中止されました"})
            break

        action_type = action_dict.get("type")
        image_name = action_dict.get("image_name")
        image_names = action_dict.get("image_names")
        text_id = action_dict.get("text_id")
        fixed_text = action_dict.get("text")  # paste_fixed用
        seconds = action_dict.get("seconds")
        count = action_dict.get("count")
        flow_name = action_dict.get("flow_name")
        group_name = action_dict.get("group_name")
        loop_count = action_dict.get("loop_count", 30)
        loop_interval = action_dict.get("loop_interval", 10)

        # エラー時に画像名を含めるためのコンテキスト情報
        action_context = ""
        if image_name:
            action_context = image_name
        elif image_names:
            action_context = " / ".join(image_names)

        try:
            if action_type == "click":
                result = execute_click(image_name, confidence, min_confidence)
            elif action_type == "click_if_exists":
                result = execute_click_if_exists(image_name, confidence, min_confidence)
            elif action_type == "click_or":
                result = execute_click_or(image_names, confidence, min_confidence)
            elif action_type == "paste":
                result = execute_paste(text_id, texts, flow_name_for_paste)
            elif action_type == "paste_fixed":
                result = execute_paste_fixed(fixed_text, flow_name_for_paste)
            elif action_type == "wait":
                result = execute_wait(image_name, confidence, wait_timeout, cursor_speed)
            elif action_type == "wait_disappear":
                result = execute_wait_disappear(image_name, confidence, wait_timeout, cursor_speed)
            elif action_type == "wait_seconds":
                result = execute_wait_seconds(seconds)
            elif action_type == "pagedown":
                result = execute_pagedown(count)
            elif action_type == "save_to_file":
                result = execute_save_to_file(text_id, flow_name, group_name, texts)
            elif action_type == "loop_click":
                result = execute_loop_click(image_name, confidence, min_confidence, loop_count, loop_interval, execution_state)
            else:
                result = {"status": "error", "message": f"不明なアクション: {action_type}"}

            execution_state.add_result(result)

            # 次のアクションまで待機（最後以外、成功時のみ）- 1秒刻みで中止チェック
            if i < len(actions) - 1 and result["status"] == "success":
                elapsed = 0
                while elapsed < interval:
                    if execution_abort_flag or execution_state.abort_flag:
                        execution_state.add_result({"status": "aborted", "message": "[中止] アクション間待機中に中止されました"})
                        execution_state.complete(False)
                        return
                    sleep_time = min(1.0, interval - elapsed)
                    time.sleep(sleep_time)
                    elapsed += sleep_time

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(f"エラー詳細: {error_detail}")
            error_msg = f"エラー: {type(e).__name__}"
            if action_context:
                error_msg += f" ({action_context})"
            tb_lines = error_detail.strip().split('\n')
            if len(tb_lines) >= 2:
                error_msg += f" [場所: {tb_lines[-2].strip()}]"
            execution_state.add_result({"status": "error", "message": error_msg})

    restore_browser_window()
    execution_state.finish()


@app.post("/api/execute")
async def execute_actions(request: ExecuteRequest):
    """アクションを順番に実行する（バックグラウンド）"""
    global execution_abort_flag
    execution_abort_flag = False  # 実行開始時にリセット

    if not request.actions:
        raise HTTPException(status_code=400, detail="アクションが選択されていません")

    # 既に実行中かチェック
    execution_id = execution_state.start(len(request.actions))
    if execution_id is None:
        raise HTTPException(status_code=409, detail="既に実行中です。完了を待つか中止してください。")

    # リクエストを辞書に変換（スレッドに渡すため）
    request_dict = {
        "actions": [a.model_dump() for a in request.actions],
        "confidence": request.confidence,
        "min_confidence": request.min_confidence,
        "wait_timeout": request.wait_timeout,
        "cursor_speed": request.cursor_speed,
        "interval": request.interval,
        "start_delay": request.start_delay
    }

    # バックグラウンドスレッドで実行開始
    thread = threading.Thread(target=run_actions_in_background, args=(request_dict,))
    thread.start()

    # すぐにレスポンスを返す
    return {"status": "started", "execution_id": execution_id, "message": "実行を開始しました"}


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


def execute_click(image_name: str, confidence: float, min_confidence: float = 0.7, max_retries: int = 2) -> dict:
    """画像をクリック（段階的に認識精度を下げて試行、リトライ付き）"""
    if not image_name:
        return {"status": "error", "message": "画像が指定されていません"}

    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        return {"status": "error", "message": f"画像ファイルが見つかりません: {image_name}"}

    for retry in range(max_retries):
        # 段階的に精度を下げて試す（浮動小数点誤差対策で0.001の余裕）
        current_conf = confidence
        while current_conf >= min_confidence - 0.001:
            try:
                location = pyautogui.locateCenterOnScreen(str(image_path), confidence=current_conf)
                if location is not None:
                    pyautogui.click(location)
                    retry_note = f", リトライ{retry+1}回目" if retry > 0 else ""
                    if current_conf < confidence - 0.001:
                        return {"status": "success", "message": f"[クリック] {image_name} (位置: {location}, 精度{int(round(current_conf*100))}%で検出{retry_note})"}
                    else:
                        return {"status": "success", "message": f"[クリック] {image_name} (位置: {location}{retry_note})"}
            except pyautogui.ImageNotFoundException:
                pass
            except Exception:
                pass
            current_conf -= 0.02

        # 見つからなかった場合、次のリトライ前に1秒待機
        if retry < max_retries - 1:
            time.sleep(1.0)

    # 最低精度でも見つからなかった場合、どの程度で検出可能か調べる
    found_conf, _ = find_best_match_confidence(str(image_path), min_confidence)

    if found_conf is not None:
        return {"status": "not_found", "message": f"画面上に画像が見つかりません: {image_name} ({max_retries}回試行後も失敗、最大{int(found_conf*100)}%で検出、最低設定は{int(min_confidence*100)}%)"}

    return {"status": "not_found", "message": f"画面上に画像が見つかりません: {image_name} ({max_retries}回試行後も失敗、30%未満)"}


def execute_click_if_exists(image_name: str, confidence: float, min_confidence: float = 0.7) -> dict:
    """画像が存在すればクリック、なければスキップ（エラーにしない）"""
    if not image_name:
        return {"status": "skipped", "message": "[スキップ] 画像が指定されていません"}

    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        return {"status": "skipped", "message": f"[スキップ] 画像ファイルが見つかりません: {image_name}"}

    # 段階的に精度を下げて試す
    current_conf = confidence
    while current_conf >= min_confidence - 0.001:
        try:
            location = pyautogui.locateCenterOnScreen(str(image_path), confidence=current_conf)
            if location is not None:
                pyautogui.click(location)
                if current_conf < confidence - 0.001:
                    return {"status": "success", "message": f"[条件クリック] {image_name} (位置: {location}, 精度{int(round(current_conf*100))}%で検出)"}
                else:
                    return {"status": "success", "message": f"[条件クリック] {image_name} (位置: {location})"}
        except pyautogui.ImageNotFoundException:
            pass
        except Exception:
            pass
        current_conf -= 0.02

    # 見つからなかった場合はスキップ（エラーではない）
    return {"status": "skipped", "message": f"[スキップ] 画像が見つからないためスキップ: {image_name}"}


def execute_click_or(image_names: list[str], confidence: float, min_confidence: float = 0.7, max_retries: int = 2) -> dict:
    """複数画像のうち最初に見つかった画像をクリック（段階的に精度を下げて試行、リトライ付き）"""
    if not image_names or len(image_names) == 0:
        return {"status": "error", "message": "画像が指定されていません"}

    for retry in range(max_retries):
        # 各精度レベルで全画像を試す
        current_conf = confidence
        while current_conf >= min_confidence - 0.001:
            for image_name in image_names:
                image_path = IMAGES_DIR / image_name
                if not image_path.exists():
                    continue

                try:
                    location = pyautogui.locateCenterOnScreen(str(image_path), confidence=current_conf)
                    if location is not None:
                        pyautogui.click(location)
                        retry_note = f", リトライ{retry+1}回目" if retry > 0 else ""
                        if current_conf < confidence - 0.001:
                            return {"status": "success", "message": f"[クリックOR] {image_name} (位置: {location}, 精度{int(current_conf*100)}%で検出{retry_note})"}
                        else:
                            return {"status": "success", "message": f"[クリックOR] {image_name} (位置: {location}{retry_note})"}
                except pyautogui.ImageNotFoundException:
                    pass
                except Exception:
                    pass

            current_conf -= 0.02

        # 見つからなかった場合、次のリトライ前に1秒待機
        if retry < max_retries - 1:
            time.sleep(1.0)

    # どの画像も見つからなかった場合、検出可能な精度を調べる
    not_found_details = []
    for image_name in image_names:
        image_path = IMAGES_DIR / image_name
        if not image_path.exists():
            not_found_details.append(f"{image_name}(ファイルなし)")
            continue

        found_conf, _ = find_best_match_confidence(str(image_path), min_confidence)
        if found_conf is not None:
            not_found_details.append(f"{image_name}({int(found_conf*100)}%)")
        else:
            not_found_details.append(f"{image_name}(30%未満)")

    detail_str = " / ".join(not_found_details)
    return {"status": "not_found", "message": f"画面上にどの画像も見つかりません: {detail_str} ({max_retries}回試行後も失敗、最低設定: {int(min_confidence*100)}%)"}


def execute_paste(text_id: str, texts: dict, flow_name: str = None) -> dict:
    """テキストを貼り付け（IDベース）"""
    if text_id is None or text_id not in texts:
        return {"status": "error", "message": f"テキストが見つかりません: ID={text_id}"}

    text = texts[text_id]["text"]

    # Grok系フローの場合、日本語で回答するよう指示を追加
    print(f"[DEBUG] execute_paste: flow_name={flow_name}")
    if flow_name and ("Grok" in flow_name or "grok" in flow_name):
        text = text + "\n\n※日本語で回答してください。"
        print(f"[DEBUG] Grok detected, added Japanese instruction")

    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')
    return {"status": "success", "message": f"[貼付] [ID:{text_id}] {text[:30]}..."}


def execute_paste_fixed(text: str, flow_name: str = None) -> dict:
    """固定テキストを貼り付け（フロー専用、テキストID不要）"""
    if not text:
        return {"status": "error", "message": "固定テキストが指定されていません"}

    paste_text = text

    # Grok系フローの場合、日本語で回答するよう指示を追加
    if flow_name and ("Grok" in flow_name or "grok" in flow_name):
        paste_text = paste_text + "\n\n※日本語で回答してください。"

    pyperclip.copy(paste_text)
    pyautogui.hotkey('ctrl', 'v')
    display_text = text[:30] + "..." if len(text) > 30 else text
    return {"status": "success", "message": f"[固定貼付] {display_text}"}


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
        # 中止チェック（両方のフラグをチェック）
        if execution_abort_flag or execution_state.abort_flag:
            return {"status": "aborted", "message": f"[待機] 中止されました: {image_name}"}

        try:
            location = pyautogui.locateCenterOnScreen(str(image_path), confidence=confidence)
        except pyautogui.ImageNotFoundException:
            location = None
        except Exception:
            location = None

        if location is not None:
            # 検出した画像の100ピクセル上にカーソルを移動（スクロールエリアをクリックしやすくする）
            target_y = max(0, location.y - 100)
            pyautogui.moveTo(location.x, target_y, duration=0.2)
            return {"status": "success", "message": f"[待機] 画像を検出: {image_name} (位置: {location}, カーソル移動先: y={target_y})"}

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


def execute_wait_disappear(image_name: str, confidence: float, timeout: float, cursor_speed: float = 0.5) -> dict:
    """画像が消えるまで待機（消失待機）

    誤検知防止のため、連続3回「見つからない」を確認してから成功とする
    """
    if not image_name:
        return {"status": "error", "message": "画像が指定されていません"}

    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        return {"status": "error", "message": f"画像ファイルが見つかりません: {image_name}"}

    # まず画像が存在することを確認
    try:
        location = pyautogui.locateCenterOnScreen(str(image_path), confidence=confidence)
    except Exception as e:
        print(f"[DEBUG] 消失待機 初回チェック例外: {type(e).__name__}: {e}")
        location = None

    if location is None:
        return {"status": "success", "message": f"[消失待機] 画像は既に画面にありません: {image_name}"}

    start_time = time.time()
    move_direction = 1
    move_amount = 100
    consecutive_not_found = 0  # 連続して見つからなかった回数
    required_consecutive = 3   # 成功判定に必要な連続回数
    check_count = 0            # 総チェック回数（ログ用）

    while time.time() - start_time < timeout:
        # 中止チェック（両方のフラグをチェック）
        if execution_abort_flag or execution_state.abort_flag:
            return {"status": "aborted", "message": f"[消失待機] 中止されました: {image_name}"}

        check_count += 1

        # 指定された信頼度で検索
        try:
            location = pyautogui.locateCenterOnScreen(str(image_path), confidence=confidence)
        except pyautogui.ImageNotFoundException:
            location = None
        except Exception:
            location = None

        # 見つからなかった場合、低い信頼度で再チェックして実際の検出状況を確認
        actual_conf = None
        if location is None:
            for test_conf in [0.3, 0.5, 0.7, 0.8, 0.9]:
                try:
                    test_loc = pyautogui.locateCenterOnScreen(str(image_path), confidence=test_conf)
                    if test_loc:
                        actual_conf = test_conf
                        break
                except:
                    pass

        if location is None:
            consecutive_not_found += 1
            if actual_conf:
                print(f"[DEBUG] 消失待機 チェック#{check_count}: 閾値{int(confidence*100)}%で不検出、実際は{int(actual_conf*100)}%で存在 ({consecutive_not_found}/{required_consecutive}回連続)")
            else:
                print(f"[DEBUG] 消失待機 チェック#{check_count}: 完全に見つからず（30%未満） ({consecutive_not_found}/{required_consecutive}回連続)")

            # 50%以上で見つかる場合は「まだ存在している」のでリセット
            if actual_conf and actual_conf >= 0.5:
                print(f"[DEBUG] → 50%以上で検出されるため、連続カウントをリセット")
                consecutive_not_found = 0
            elif consecutive_not_found >= required_consecutive:
                elapsed = time.time() - start_time
                return {"status": "success", "message": f"[消失待機] 画像が消えました: {image_name} ({elapsed:.1f}秒後、{check_count}回チェック、{required_consecutive}回連続不検出で確定)"}
        else:
            if consecutive_not_found > 0:
                print(f"[DEBUG] 消失待機 チェック#{check_count}: 検出 (連続不検出{consecutive_not_found}回→リセット)")
            consecutive_not_found = 0  # 見つかったらリセット
        # 待機中を示すためにカーソルを左右にスムーズに動かす
        current_pos = pyautogui.position()
        target_x = current_pos[0] + (move_amount * move_direction)
        smooth_move_cursor(target_x, current_pos[1], cursor_speed)
        move_direction *= -1

        time.sleep(0.5)  # 画像チェックの間隔（消失待機は少し長めに）

    return {"status": "timeout", "message": f"タイムアウト: {image_name} が {timeout}秒以内に消えませんでした ({check_count}回チェック)"}



def execute_wait_seconds(seconds: float) -> dict:
    """指定秒数だけ待機（秒数待機）"""
    global execution_abort_flag
    if seconds is None or seconds < 0:
        return {"status": "error", "message": "秒数が指定されていません"}

    # 1秒刻みで中止チェック
    elapsed = 0
    while elapsed < seconds:
        if execution_abort_flag or execution_state.abort_flag:
            return {"status": "aborted", "message": f"[秒数待機] {elapsed}/{seconds}秒で中止されました"}
        sleep_time = min(1.0, seconds - elapsed)
        time.sleep(sleep_time)
        elapsed += sleep_time
    return {"status": "success", "message": f"[秒数待機] {seconds}秒待機しました"}


def execute_pagedown(count: int) -> dict:
    """PageDownキーを指定回数押す（スクロール）"""
    if count is None or count < 1:
        count = 1

    # まず現在位置でクリックしてフォーカスを当てる
    click_pos = pyautogui.position()
    pyautogui.click()
    time.sleep(0.1)

    for i in range(count):
        pyautogui.press('pagedown')
        if i < count - 1:
            time.sleep(0.1)  # 連打時の間隔

    return {"status": "success", "message": f"[PageDown] {count}回押しました (クリック位置: {click_pos})"}


def sanitize_filename(text: str, max_length: int = 30) -> str:
    """ファイル名に使えない文字を除去し、長さを制限"""
    import re
    # ファイル名に使えない文字を除去
    sanitized = re.sub(r'[\\/:*?"<>|\r\n\t]', '', text)
    # 空白を_に置換
    sanitized = re.sub(r'\s+', '_', sanitized)
    # 長さ制限
    if len(sanitized) > max_length:
        sanitized = sanitized[:max_length]
    return sanitized or "untitled"

def analyze_file_character_counts(filepath: Path) -> dict:
    """ファイル内の各AIセクションの文字数をカウント"""
    import re

    if not filepath.exists():
        return {"total": 0, "sections": []}

    try:
        with open(filepath, "r", encoding="utf-8") as f:
            file_content = f.read()
    except Exception:
        return {"total": 0, "sections": []}

    # セパレータで分割
    # ファイル形式:
    # ##################################################
    # フロー: name
    # 日時: timestamp
    # ##################################################
    # [body content]
    separator = "#" * 50
    sections = file_content.split(separator)

    results = []
    i = 0
    while i < len(sections):
        section = sections[i]
        if "フロー:" in section:
            # フロー名を抽出
            match = re.search(r"フロー:\s*(.+)", section)
            if match:
                flow_name = match.group(1).strip()
                # 次のセクションが本文
                body_text = ""
                if i + 1 < len(sections):
                    body_text = sections[i + 1].strip()
                char_count = len(body_text)
                line_count = len([l for l in body_text.split("\n") if l.strip()])
                results.append({
                    "flow_name": flow_name,
                    "chars": char_count,
                    "lines": line_count
                })
        i += 1

    total_chars = sum(r["chars"] for r in results)
    return {"total": total_chars, "sections": results}



def execute_save_to_file(text_id: str, flow_name: str, group_name: str, texts: dict) -> dict:
    """クリップボードの内容をファイルに追記保存"""
    from datetime import datetime
    import re

    # グループ名の日本語ラベル
    GROUP_LABELS = {
        'ai-normal': 'AI-通常',
        'ai-dr': 'AI-DR',
        'ai-chat': 'AI会話',
        'blog': 'ブログ投稿',
        'image-gen': '画像生成'
    }

    # テキストIDからファイル名を決定
    if text_id and text_id in texts:
        text_content = texts[text_id]["text"]
        text_part = sanitize_filename(text_content, 30)
    else:
        text_content = None
        text_part = f"text_{text_id or 'unknown'}"

    # グループ名をファイル名に使う（グループがない場合は「未分類」）
    group_label = GROUP_LABELS.get(group_name, '未分類') if group_name else '未分類'
    filename_base = f"{group_label}_{text_part}"

    # 日付フォルダを作成
    today = datetime.now().strftime("%Y-%m-%d")
    output_dir = Path(__file__).parent / "output" / today
    output_dir.mkdir(parents=True, exist_ok=True)

    # ファイルパス
    filepath = output_dir / f"{filename_base}.txt"

    # クリップボードから内容を取得
    try:
        clipboard_content = pyperclip.paste()
    except Exception as e:
        return {"status": "error", "message": f"クリップボード取得エラー: {e}"}

    # クリップボードが空かチェック
    if not clipboard_content or not clipboard_content.strip():
        return {"status": "error", "message": f"[ファイル保存] クリップボードが空です。コピー操作を忘れていませんか？"}

    # クリップボードの内容がペーストしたテキストと同じ場合は警告
    if text_content and clipboard_content.strip() == text_content.strip():
        return {"status": "error", "message": f"[ファイル保存] クリップボードの内容がプロンプトと同じです。回答をコピーし忘れていませんか？"}

    # 保存内容を作成
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    separator = "#" * 50
    save_content = f"\n{separator}\nフロー: {flow_name or '(名前なし)'}\n日時: {timestamp}\n{separator}\n{clipboard_content}\n"

    # 今回保存する内容の文字数をカウント（clipboard_contentの文字数）
    content_char_count = len(clipboard_content.strip())

    # ファイルに追記
    try:
        with open(filepath, "a", encoding="utf-8") as f:
            f.write(save_content)
            # AI-通常、AI-DRの場合は文字数ログも追記
            if group_name in ['ai-normal', 'ai-dr']:
                char_log = f"\n---\n📊 {timestamp} | {flow_name}: {content_char_count}文字\n"
                f.write(char_log)

        # 文字数カウントを取得（AI-通常、AI-DRの場合）
        char_stats = None
        if group_name in ['ai-normal', 'ai-dr']:
            char_stats = analyze_file_character_counts(filepath)

        result = {"status": "success", "message": f"[ファイル保存] {filepath.name} に追記しました ({content_char_count}文字)"}
        if char_stats:
            result["char_stats"] = char_stats
        return result
    except Exception as e:
        return {"status": "error", "message": f"ファイル保存エラー: {e}"}


def execute_loop_click(image_name: str, confidence: float, min_confidence: float, loop_count: int, loop_interval: int, execution_state) -> dict:
    """画像を指定回数、指定間隔でループクリック"""
    global execution_abort_flag

    if not image_name:
        return {"status": "error", "message": "画像が指定されていません"}

    image_path = IMAGES_DIR / image_name
    if not image_path.exists():
        return {"status": "error", "message": f"画像ファイルが見つかりません: {image_name}"}

    success_count = 0
    fail_count = 0

    for i in range(loop_count):
        # 中止チェック
        if execution_state.abort_flag or execution_abort_flag:
            return {"status": "aborted", "message": f"[ループクリック] {i}/{loop_count}回で中止 (成功: {success_count}, 失敗: {fail_count})"}

        # クリック試行
        clicked = False
        current_conf = confidence
        while current_conf >= min_confidence - 0.001:
            try:
                location = pyautogui.locateCenterOnScreen(str(image_path), confidence=current_conf)
                if location is not None:
                    pyautogui.click(location)
                    clicked = True
                    success_count += 1
                    break
            except:
                pass
            current_conf -= 0.02

        if not clicked:
            fail_count += 1

        # 進捗を報告（10回ごと、または最初と最後）
        if i == 0 or (i + 1) % 10 == 0 or i == loop_count - 1:
            execution_state.add_result({
                "status": "info",
                "message": f"[ループクリック] {i + 1}/{loop_count}回完了 (成功: {success_count}, 失敗: {fail_count})"
            })

        # 最後のループ以外は間隔待機 - 1秒刻みで中止チェック
        if i < loop_count - 1:
            elapsed = 0
            while elapsed < loop_interval:
                if execution_abort_flag or execution_state.abort_flag:
                    return {"status": "aborted", "message": f"[ループクリック] {i+1}/{loop_count}回で中止 (成功: {success_count}, 失敗: {fail_count})"}
                sleep_time = min(1.0, loop_interval - elapsed)
                time.sleep(sleep_time)
                elapsed += sleep_time

    return {"status": "success", "message": f"[ループクリック] {loop_count}回完了 (成功: {success_count}, 失敗: {fail_count})"}


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
