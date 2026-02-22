"""Flask backend for DeepExtract - 文档结构化萃取服务"""

import os
import sys
import uuid
import json
import time
import shutil
import threading
from pathlib import Path
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import tempfile

# Add parent dir to path for imports
SCRIPT_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(SCRIPT_DIR))

import mineru_extract
import md2word_final

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "deepextract"
UPLOAD_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {
    "pdf",
    "doc",
    "docx",
    "ppt",
    "pptx",
    "png",
    "jpg",
    "jpeg",
    "html",
    "md",
    "markdown",
}
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# In-memory task store
tasks = {}
tasks_lock = threading.Lock()

TASK_RETENTION_SECONDS = 5 * 60
TASK_CLEANUP_INTERVAL_SECONDS = 30


def get_extension(filename):
    return filename.rsplit(".", 1)[1].lower() if "." in filename else ""


def allowed_file(filename):
    return "." in filename and get_extension(filename) in ALLOWED_EXTENSIONS


def _now():
    return time.time()


def _mark_task_expire(task):
    task["completed_at"] = _now()
    task["expires_at"] = task["completed_at"] + TASK_RETENTION_SECONDS


def _remove_task_artifacts(task):
    task_dir = task.get("task_dir")
    if task_dir and os.path.exists(task_dir):
        shutil.rmtree(task_dir, ignore_errors=True)


def _cleanup_expired_tasks(now_ts=None):
    current_ts = _now() if now_ts is None else now_ts
    expired = []

    with tasks_lock:
        for task_id, task in tasks.items():
            expires_at = task.get("expires_at")
            if expires_at and expires_at <= current_ts:
                expired.append((task_id, task.get("task_dir")))

    if not expired:
        return 0

    for _, task_dir in expired:
        if task_dir and os.path.exists(task_dir):
            shutil.rmtree(task_dir, ignore_errors=True)

    with tasks_lock:
        for task_id, _ in expired:
            tasks.pop(task_id, None)

    return len(expired)


def _cleanup_loop():
    while True:
        try:
            _cleanup_expired_tasks()
        except Exception as exc:
            print(f"cleanup worker error: {exc}")
        time.sleep(TASK_CLEANUP_INTERVAL_SECONDS)


def process_task(task_id, input_path, target_format, original_name, docx_options=None):
    """Background task: convert file to target format."""
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        return

    try:
        ext = get_extension(original_name)
        stem = Path(original_name).stem
        task_dir = str(Path(input_path).parent)

        if ext in ("md", "markdown"):
            # Input is already Markdown
            if target_format == "md":
                # 直接返回 md 文件
                task["result_path"] = input_path
                task["result_name"] = original_name
            else:
                # Markdown → Word（直接转换，无需 MinerU）
                task["status_text"] = "正在转换为 Word 文档"
                task["progress"] = 30
                output_path = os.path.join(task_dir, f"{stem}.docx")
                md2word_final.convert_with_python_docx(
                    input_path,
                    output_path,
                    formula_numbering_mode=(docx_options or {}).get("formula_numbering_mode"),
                    doc_style_options=docx_options,
                )
                task["result_path"] = output_path
                task["result_name"] = f"{stem}.docx"
        else:
            # 其他格式 → 先走 MinerU API
            task["status_text"] = "正在上传至解析引擎"
            task["progress"] = 10

            task["status_text"] = "正在深度解析文档结构"
            task["progress"] = 20
            result = mineru_extract.upload_and_extract(input_path)
            task["progress"] = 70

            if target_format == "md":
                # 导出 Markdown → 返回 ZIP 压缩包（含图片）
                task["result_path"] = result["zip_path"]
                task["result_name"] = f"{stem}.zip"
            else:
                # 导出 Word → 用解压目录中的 md + 图片生成 docx
                task["status_text"] = "正在生成 Word 文档"
                task["progress"] = 80
                output_path = os.path.join(task_dir, f"{stem}.docx")
                md2word_final.convert_with_python_docx(
                    result["md_path"],
                    output_path,
                    formula_numbering_mode=(docx_options or {}).get("formula_numbering_mode"),
                    doc_style_options=docx_options,
                )
                task["result_path"] = output_path
                task["result_name"] = f"{stem}.docx"

        task["state"] = "done"
        task["progress"] = 100
        _mark_task_expire(task)

    except Exception as e:
        task["state"] = "failed"
        task["error"] = str(e)
        _mark_task_expire(task)
        print(f"Task {task_id} failed: {e}")


@app.route("/api/upload", methods=["POST"])
def upload():
    _cleanup_expired_tasks()

    if "file" not in request.files:
        return jsonify({"error": "没有上传文件"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "未选择文件"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "不支持的文件格式"}), 400

    target = request.form.get("target", "md")
    if target not in ("md", "doc"):
        return jsonify({"error": "不支持的目标格式"}), 400

    docx_options = {}
    raw_docx_options = request.form.get("docx_options", "").strip()
    if raw_docx_options:
        try:
            parsed = json.loads(raw_docx_options)
            if isinstance(parsed, dict):
                docx_options = parsed
        except json.JSONDecodeError:
            return jsonify({"error": "Word 配置格式错误"}), 400

    # Create task directory
    task_id = str(uuid.uuid4())[:8]
    task_dir = UPLOAD_FOLDER / task_id
    task_dir.mkdir(exist_ok=True)

    # Save file (handle Chinese filenames)
    original_name = file.filename
    safe_name = secure_filename(file.filename)
    if not safe_name:
        safe_name = f"upload.{get_extension(original_name)}"
    input_path = str(task_dir / safe_name)
    file.save(input_path)

    with tasks_lock:
        tasks[task_id] = {
            "state": "processing",
            "progress": 0,
            "status_text": "准备中",
            "result_path": None,
            "result_name": None,
            "error": None,
            "task_dir": str(task_dir),
            "created_at": _now(),
            "completed_at": None,
            "expires_at": None,
        }

    threading.Thread(
        target=process_task,
        args=(task_id, input_path, target, original_name, docx_options),
        daemon=True,
    ).start()

    return jsonify({"task_id": task_id})


@app.route("/api/status/<task_id>", methods=["GET"])
def status(task_id):
    _cleanup_expired_tasks()
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在"}), 404

    return jsonify(
        {
            "state": task["state"],
            "progress": task["progress"],
            "status_text": task.get("status_text", ""),
            "error": task.get("error"),
            "result_name": task.get("result_name"),
        }
    )


@app.route("/api/download/<task_id>", methods=["GET"])
def download(task_id):
    _cleanup_expired_tasks()
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在"}), 404
    if task["state"] != "done":
        return jsonify({"error": "任务尚未完成"}), 400

    result_path = task["result_path"]
    if not result_path or not os.path.exists(result_path):
        return jsonify({"error": "结果文件不存在"}), 404

    return send_file(result_path, as_attachment=True, download_name=task["result_name"])


cleanup_thread = threading.Thread(target=_cleanup_loop, daemon=True)
cleanup_thread.start()


@app.route("/")
def index():
    """Serve the frontend HTML file."""
    frontend_path = SCRIPT_DIR / "front.html"
    if frontend_path.exists():
        return send_file(str(frontend_path))
    return jsonify({"error": "Frontend not found"}), 404


@app.route("/<path:filename>")
def static_files(filename):
    """Serve static files (images, css, js, etc.)."""
    file_path = SCRIPT_DIR / filename
    if file_path.exists() and file_path.is_file():
        return send_file(str(file_path))
    return jsonify({"error": "File not found"}), 404


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    print("=" * 50)
    print("  DeepExtract - 文档结构化萃取服务")
    print("=" * 50)
    print(f"  地址: http://localhost:5000")
    print("=" * 50)
    app.run(debug=True, host="0.0.0.0", port=5000)
