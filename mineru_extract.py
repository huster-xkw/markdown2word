"""MinerU API: 任意格式文件 → Markdown（→ Word）"""
import os
import sys
import time
import zipfile
import requests
from pathlib import Path
from typing import Optional

# API 配置
API_BASE = "https://mineru.net"


def _get_api_token() -> str:
    """Read MinerU API key from env or local apikey.md."""
    token = os.getenv("MINERU_API_KEY", "").strip()
    if token:
        return token

    key_file = Path(__file__).parent / "apikey.md"
    if key_file.exists():
        for line in key_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("MINERU_API_KEY="):
                return line.split("=", 1)[1].strip()
    return ""


def _get_headers() -> dict:
    token = _get_api_token()
    if not token:
        raise RuntimeError(
            "未配置 MINERU_API_KEY。请在环境变量中设置，或在 apikey.md 中填写 MINERU_API_KEY=你的Key"
        )
    return {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}",
    }

POLL_INTERVAL = 5  # 轮询间隔（秒）
POLL_TIMEOUT = 300  # 最大等待时间（秒）


def upload_and_extract(file_path: str, model_version: str = "vlm") -> dict:
    """上传本地文件到 MinerU 并获取解析后的 Markdown 内容。

    Args:
        file_path: 本地文件路径
        model_version: 模型版本，默认 vlm

    Returns:
        dict: {
            'md_path': markdown 文件路径,
            'zip_path': 原始 zip 文件路径,
            'extract_dir': 解压目录路径,
        }
    """
    input_path = Path(file_path)
    if not input_path.exists():
        raise FileNotFoundError(f"文件不存在: {input_path}")

    file_name = input_path.name
    print(f"[1/4] 申请上传 URL ... ({file_name})")

    # 1. 申请上传 URL
    resp = requests.post(
        f"{API_BASE}/api/v4/file-urls/batch",
        headers=_get_headers(),
        json={
            "files": [{"name": file_name}],
            "model_version": model_version,
            "enable_formula": True,
            "enable_table": True,
            "language": "ch",
        },
    )
    resp.raise_for_status()
    result = resp.json()
    if result["code"] != 0:
        raise RuntimeError(f"申请上传 URL 失败: {result.get('msg')}")

    batch_id = result["data"]["batch_id"]
    upload_url = result["data"]["file_urls"][0]
    print(f"  batch_id: {batch_id}")

    # 2. PUT 上传文件
    print(f"[2/4] 上传文件 ...")
    with open(input_path, "rb") as f:
        resp_upload = requests.put(upload_url, data=f)
    if resp_upload.status_code != 200:
        raise RuntimeError(f"文件上传失败: HTTP {resp_upload.status_code}")
    print(f"  上传成功")

    # 3. 轮询任务结果
    print(f"[3/4] 等待解析完成 ...")
    start = time.time()
    while True:
        elapsed = time.time() - start
        if elapsed > POLL_TIMEOUT:
            raise TimeoutError(f"解析超时（已等待 {POLL_TIMEOUT}s）")

        resp_result = requests.get(
            f"{API_BASE}/api/v4/extract-results/batch/{batch_id}",
            headers=_get_headers(),
        )
        resp_result.raise_for_status()
        data = resp_result.json()

        if data["code"] != 0:
            raise RuntimeError(f"查询失败: {data.get('msg')}")

        extract = data["data"]["extract_result"][0]
        state = extract["state"]

        if state == "done":
            zip_url = extract["full_zip_url"]
            print(f"  解析完成！耗时 {elapsed:.1f}s")
            break
        elif state == "failed":
            raise RuntimeError(f"解析失败: {extract.get('err_msg')}")
        else:
            progress = extract.get("extract_progress", {})
            extracted = progress.get("extracted_pages", "?")
            total = progress.get("total_pages", "?")
            print(f"  状态: {state}  进度: {extracted}/{total}  已等待 {elapsed:.0f}s")
            time.sleep(POLL_INTERVAL)

    # 4. 下载并解压结果（带重试，阿里云 OSS 偶尔会断连）
    print(f"[4/4] 下载解析结果 ...")
    output_dir = input_path.parent / f"{input_path.stem}_mineru_output"
    output_dir.mkdir(exist_ok=True)

    zip_path = output_dir / "result.zip"
    max_retries = 3
    # 绕过系统代理直连 OSS，避免代理导致下载失败
    no_proxy = {"http": "", "https": ""}
    for attempt in range(1, max_retries + 1):
        try:
            resp_zip = requests.get(zip_url, timeout=120, proxies=no_proxy)
            resp_zip.raise_for_status()
            zip_path.write_bytes(resp_zip.content)
            break
        except (requests.exceptions.SSLError,
                requests.exceptions.ConnectionError,
                requests.exceptions.ProxyError) as e:
            if attempt < max_retries:
                wait = attempt * 3
                print(f"  下载失败（第{attempt}次），{wait}s 后重试: {e}")
                time.sleep(wait)
            else:
                raise RuntimeError(f"下载结果 ZIP 失败（已重试{max_retries}次）: {e}")

    # 解压
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(output_dir)
    print(f"  解压到: {output_dir}")

    # 查找 markdown 文件
    md_files = list(output_dir.rglob("*.md"))
    if not md_files:
        raise FileNotFoundError(f"解压结果中未找到 .md 文件")

    md_file = md_files[0]
    md_content = md_file.read_text(encoding="utf-8")
    print(f"  Markdown 文件: {md_file}")
    print(f"  内容长度: {len(md_content)} 字符")

    return {
        'md_path': str(md_file),
        'zip_path': str(zip_path),
        'extract_dir': str(output_dir),
    }


def extract_to_word(file_path: str, output_docx: Optional[str] = None, model_version: str = "vlm"):
    """完整流程：任意格式 → Markdown → Word"""
    from md2word_final import convert_with_python_docx

    result = upload_and_extract(file_path, model_version)
    md_path = result['md_path']

    if output_docx is None:
        output_docx = str(Path(file_path).with_suffix(".docx"))

    print(f"\n[转换] Markdown → Word ...")
    convert_with_python_docx(md_path, output_docx)
    print(f"最终输出: {output_docx}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法:")
        print("  提取为 Markdown:  python mineru_extract.py <文件路径>")
        print("  提取为 Word:      python mineru_extract.py <文件路径> --word [输出路径]")
        sys.exit(1)

    input_file = sys.argv[1]
    to_word = "--word" in sys.argv

    if to_word:
        # 找 --word 后面的可选输出路径
        idx = sys.argv.index("--word")
        output = sys.argv[idx + 1] if idx + 1 < len(sys.argv) else None
        extract_to_word(input_file, output)
    else:
        result = upload_and_extract(input_file)
        print(f"\nMarkdown 输出: {result['md_path']}")
        print(f"ZIP 文件: {result['zip_path']}")
        print(f"解压目录: {result['extract_dir']}")
