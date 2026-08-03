#!/usr/bin/env python3
"""低成本实跑链路：完整时序扫描 -> 候选窗口密集复核 -> DeepSeek 整理 -> Excel。

车牌只在豆包对原始连续帧作出结构化判断、且确定性校验通过后进入结果。
该脚本不做官方认定，不使用生成式修复，也不提交举报。
"""

import argparse
import base64
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    sys.exit("缺少依赖，请先运行：pip install openpyxl pillow")

from build_analysis_workbook import write_workbook


VIDEO_SUFFIXES = {".mp4", ".mov", ".avi", ".mkv", ".webm", ".m4v"}
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
SUPPORTED_BEHAVIORS = [
    "远光灯炫目",
    "闯红灯",
    "逆行",
    "加塞/违规变道",
    "违停",
    "压实线/压导流线",
    "压水花溅人",
    "不礼让斑马线",
    "开车打手机",
    "占用应急车道",
    "货车违反车道通行规定",
]
PLATE_KEYS = {
    "plate",
    "plate_number",
    "plate_text",
    "license_plate",
    "车牌",
    "车牌号",
    "车牌字符",
}
PLATE_PATTERN = re.compile(
    r"[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领警学港澳]"
    r"[A-Z][·•\- ]?[A-Z0-9]{5,6}",
    re.IGNORECASE,
)
PLATE_EXACT_PATTERN = re.compile(
    r"^[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼]"
    r"[A-Z][A-Z0-9]{5,6}$"
)


@dataclass(frozen=True)
class Frame:
    frame_id: str
    source_file: str
    timestamp: float
    path: Path


def post_json(url: str, api_key: str, payload: dict, timeout: int) -> dict:
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")[:1200]
        raise RuntimeError(f"API 返回 HTTP {exc.code}: {body}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"API 连接失败: {exc.reason}") from exc


def extract_message_content(response: dict) -> str:
    try:
        content = response["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise ValueError("API 响应缺少 choices[0].message.content") from exc
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, dict) and isinstance(item.get("text"), str):
                parts.append(item["text"])
        if parts:
            return "\n".join(parts)
    raise ValueError("API 响应 content 格式不受支持")


def parse_json_content(content: str) -> dict:
    stripped = content.strip()
    if stripped.startswith("```"):
        stripped = re.sub(r"^```(?:json)?\s*", "", stripped)
        stripped = re.sub(r"\s*```$", "", stripped)
    data = json.loads(stripped)
    if not isinstance(data, dict):
        raise ValueError("模型输出必须是 JSON 对象")
    return data


def image_data_url(path: Path, max_dimension: int = 1600) -> str:
    with Image.open(path) as image:
        image = image.convert("RGB")
        image.thumbnail((max_dimension, max_dimension))
        buffer = io.BytesIO()
        image.save(buffer, format="JPEG", quality=82, optimize=True)
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def find_ffmpeg() -> str:
    configured = os.environ.get("FFMPEG_BIN")
    if configured and Path(configured).is_file():
        return configured
    found = shutil.which("ffmpeg")
    if not found:
        raise RuntimeError("处理视频需要 ffmpeg；请先安装或设置 FFMPEG_BIN")
    return found


def extract_video_frames(
    video_path: Path,
    work_dir: Path,
    fps: float,
    max_frames: int,
) -> list[Frame]:
    ffmpeg = find_ffmpeg()
    prefix = re.sub(r"[^a-zA-Z0-9_-]+", "-", video_path.stem)[:40] or "video"
    output_pattern = work_dir / f"{prefix}-frame-%06d.jpg"
    command = [
        ffmpeg,
        "-hide_banner",
        "-loglevel",
        "error",
        "-i",
        str(video_path),
        "-vf",
        f"fps={fps}",
        "-frames:v",
        str(max_frames),
        "-q:v",
        "3",
        str(output_pattern),
    ]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"ffmpeg 抽帧失败: {result.stderr.strip()[:1200]}")
    paths = sorted(work_dir.glob(f"{prefix}-frame-*.jpg"))
    if not paths:
        raise RuntimeError(f"视频未抽取到画面: {video_path}")
    return [
        Frame(
            frame_id=f"{prefix}-f{index:06d}",
            source_file=video_path.name,
            timestamp=(index - 1) / fps,
            path=path,
        )
        for index, path in enumerate(paths, start=1)
    ]


def extract_video_window_frames(
    video_path: Path,
    work_dir: Path,
    fps: float,
    start_seconds: float,
    duration_seconds: float,
    max_frames: int,
    window_id: str,
) -> list[Frame]:
    """从候选窗口密集抽取原始画面，保留绝对时间戳。"""
    ffmpeg = find_ffmpeg()
    safe_stem = re.sub(r"[^a-zA-Z0-9_-]+", "-", video_path.stem)[:32] or "video"
    safe_window = re.sub(r"[^a-zA-Z0-9_-]+", "-", window_id)[:24] or "review"
    prefix = f"{safe_stem}-{safe_window}"
    output_pattern = work_dir / f"{prefix}-frame-%06d.jpg"
    command = [
        ffmpeg,
        "-hide_banner",
        "-loglevel",
        "error",
        "-ss",
        f"{max(0.0, start_seconds):.3f}",
        "-i",
        str(video_path),
        "-t",
        f"{max(0.1, duration_seconds):.3f}",
        "-vf",
        f"fps={fps}",
        "-frames:v",
        str(max_frames),
        "-q:v",
        "2",
        str(output_pattern),
    ]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"ffmpeg 密集抽帧失败: {result.stderr.strip()[:1200]}")
    paths = sorted(work_dir.glob(f"{prefix}-frame-*.jpg"))
    return [
        Frame(
            frame_id=f"{prefix}-f{index:06d}",
            source_file=video_path.name,
            timestamp=max(0.0, start_seconds) + (index - 1) / fps,
            path=path,
        )
        for index, path in enumerate(paths, start=1)
    ]


def prepare_frames(
    inputs: list[Path],
    work_dir: Path,
    fps: float,
    max_frames: int,
) -> list[Frame]:
    frames: list[Frame] = []
    for source_index, path in enumerate(inputs, start=1):
        if not path.is_file():
            raise FileNotFoundError(f"输入文件不存在: {path}")
        suffix = path.suffix.lower()
        if suffix in VIDEO_SUFFIXES:
            remaining = max_frames - len(frames)
            if remaining <= 0:
                break
            frames.extend(extract_video_frames(path, work_dir, fps, remaining))
        elif suffix in IMAGE_SUFFIXES:
            frames.append(
                Frame(
                    frame_id=f"image-{source_index:03d}",
                    source_file=path.name,
                    timestamp=0.0,
                    path=path,
                )
            )
        else:
            raise ValueError(f"不支持的素材格式: {path}")
        if len(frames) >= max_frames:
            break
    return frames[:max_frames]


def chunks(items: list[Frame], size: int):
    for index in range(0, len(items), size):
        yield items[index : index + size]


def remove_plate_fields(value):
    """递归丢弃车牌字段，并隐去模型文本中的车牌样式字符串。"""
    if isinstance(value, dict):
        return {
            key: remove_plate_fields(item)
            for key, item in value.items()
            if str(key).lower() not in PLATE_KEYS and str(key) not in PLATE_KEYS
        }
    if isinstance(value, list):
        return [remove_plate_fields(item) for item in value]
    if isinstance(value, str):
        return PLATE_PATTERN.sub("[车牌字符已移除]", value)
    return value


def call_doubao_scan(
    frames: list[Frame],
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
) -> list[dict]:
    all_candidates: list[dict] = []
    behavior_text = "、".join(SUPPORTED_BEHAVIORS)
    for batch_index, batch in enumerate(chunks(frames, 24), start=1):
        prompt = f"""
你在执行交通证据候选扫描。图片已按时间顺序排列，每张图前有 frame_id、素材文件和时间秒数。
只列出画面支持的候选，不做最终法律认定。覆盖范围：{behavior_text}。
这是完整素材的时序扫描阶段。禁止输出或猜测任何车牌字符；只可标记 plate_visibility 为 clear/unclear/not_visible。
若单张图无法证明连续动作，必须写 evidence_status=证据不足。
输出 JSON，不要 Markdown：
{{
  "candidates": [
    {{
      "source_file": "...",
      "start_seconds": 0.0,
      "end_seconds": 0.0,
      "behavior": "...",
      "vehicle_type": "...",
      "vehicle_color": "...",
      "plate_visibility": "clear|unclear|not_visible",
      "evidence_status": "待人工复核|证据不足",
      "evidence_note": "...",
      "evidence_frame_ids": ["..."]
    }}
  ]
}}
""".strip()
        content = [{"type": "text", "text": prompt}]
        for frame in batch:
            content.append(
                {
                    "type": "text",
                    "text": (
                        f"frame_id={frame.frame_id}; source={frame.source_file}; "
                        f"seconds={frame.timestamp:.3f}"
                    ),
                }
            )
            content.append(
                {
                    "type": "image_url",
                    "image_url": {"url": image_data_url(frame.path)},
                }
            )
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": content}],
            "temperature": 0,
            "response_format": {"type": "json_object"},
        }
        response = post_json(
            f"{base_url.rstrip('/')}/chat/completions",
            api_key,
            payload,
            timeout,
        )
        parsed = remove_plate_fields(parse_json_content(extract_message_content(response)))
        candidates = parsed.get("candidates", [])
        if not isinstance(candidates, list):
            raise ValueError(f"豆包第 {batch_index} 批输出缺少 candidates 数组")
        valid_ids = {frame.frame_id for frame in batch}
        for candidate in candidates:
            if not isinstance(candidate, dict):
                continue
            candidate["evidence_frame_ids"] = [
                frame_id
                for frame_id in candidate.get("evidence_frame_ids", [])
                if frame_id in valid_ids
            ]
            all_candidates.append(candidate)
    return all_candidates


def normalize_plate_number(value: object) -> str:
    text = str(value or "").upper().strip()
    return re.sub(r"[·•\-\s]", "", text)


def accept_model_plate(
    decision: object,
    valid_frames: dict[str, Frame],
    require_consensus: bool = True,
) -> dict | None:
    """确定性车牌闸门：结构、清晰度、轨迹、帧引用和独立重读必须通过。"""
    if not isinstance(decision, dict):
        return None
    plate = normalize_plate_number(decision.get("plate_number"))
    side = str(decision.get("plate_side", "")).strip().lower()
    side_map = {"front": "前", "rear": "后", "前": "前", "后": "后"}
    if not PLATE_EXACT_PATTERN.fullmatch(plate):
        return None
    if side not in side_map:
        return None
    if decision.get("plate_fully_clear") is not True:
        return None
    if decision.get("trajectory_continuous") is not True:
        return None
    plate_ids = [
        frame_id
        for frame_id in decision.get("plate_frame_ids", [])
        if isinstance(frame_id, str) and frame_id in valid_frames
    ]
    incident_ids = [
        frame_id
        for frame_id in decision.get("incident_frame_ids", [])
        if isinstance(frame_id, str) and frame_id in valid_frames
    ]
    if not plate_ids or not incident_ids:
        return None
    source_files = {
        valid_frames[frame_id].source_file for frame_id in plate_ids + incident_ids
    }
    if len(source_files) != 1:
        return None
    consensus = decision.get("independent_consensus")
    if require_consensus:
        if not isinstance(consensus, dict) or consensus.get("passed") is not True:
            return None
        if normalize_plate_number(consensus.get("plate_number")) != plate:
            return None
        consensus_side = str(consensus.get("plate_side", "")).strip().lower()
        if side_map.get(consensus_side) != side_map[side]:
            return None
        consensus_ids = list(
            dict.fromkeys(
                frame_id
                for frame_id in consensus.get("frame_ids", [])
                if isinstance(frame_id, str) and frame_id in plate_ids
            )
        )
        source_suffix = Path(valid_frames[plate_ids[0]].source_file).suffix.lower()
        required_reads = 1 if source_suffix in IMAGE_SUFFIXES else 2
        if len(consensus_ids) < required_reads:
            return None
    return {
        "plate_number": plate,
        "plate_side": side_map[side],
        "plate_fully_clear": True,
        "trajectory_continuous": True,
        "plate_frame_ids": plate_ids[:3],
        "incident_frame_ids": incident_ids[:3],
        "independent_consensus": consensus if require_consensus else None,
    }


def call_doubao_plate_consensus(
    decision: dict,
    valid_frames: dict[str, Frame],
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
) -> dict | None:
    """不提供候选号码，让模型对原始车牌帧逐帧独立重读。"""
    frame_ids = list(dict.fromkeys(decision.get("plate_frame_ids", [])))
    frames = [valid_frames[frame_id] for frame_id in frame_ids if frame_id in valid_frames]
    if not frames:
        return None
    prompt = """
你只做车牌字符的独立复核。不要推测、补全或生成式修复。
逐帧读取前牌或后牌；任一字符不清晰时，该帧 plate_number 必须为空。
只有多张视频帧（单张图片则为该图）的完整读数逐字一致，才输出 consistent_plate_number。
输出 JSON，不要 Markdown：
{
  "reads": [
    {"frame_id": "...", "plate_number": "", "plate_side": "front|rear|", "fully_clear": false}
  ],
  "consistent_plate_number": "",
  "plate_side": "front|rear|",
  "all_characters_clear": false
}
""".strip()
    content = [{"type": "text", "text": prompt}]
    for frame in frames:
        content.append(
            {
                "type": "text",
                "text": f"frame_id={frame.frame_id}; source={frame.source_file}; seconds={frame.timestamp:.3f}",
            }
        )
        content.append(
            {"type": "image_url", "image_url": {"url": image_data_url(frame.path)}}
        )
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": content}],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    response = post_json(
        f"{base_url.rstrip('/')}/chat/completions", api_key, payload, timeout
    )
    parsed = parse_json_content(extract_message_content(response))
    if parsed.get("all_characters_clear") is not True:
        return None
    plate = normalize_plate_number(parsed.get("consistent_plate_number"))
    side = str(parsed.get("plate_side", "")).strip().lower()
    if plate != decision.get("plate_number") or side not in {"front", "rear"}:
        return None
    expected_side = "front" if decision.get("plate_side") == "前" else "rear"
    if side != expected_side:
        return None
    reads = parsed.get("reads", [])
    if not isinstance(reads, list):
        return None
    readable_ids = []
    for read in reads:
        if not isinstance(read, dict) or read.get("fully_clear") is not True:
            continue
        frame_id = read.get("frame_id")
        read_plate = normalize_plate_number(read.get("plate_number"))
        read_side = str(read.get("plate_side", "")).strip().lower()
        if frame_id in frame_ids and read_plate == plate and read_side == side:
            readable_ids.append(frame_id)
    readable_ids = list(dict.fromkeys(readable_ids))
    source_suffix = Path(frames[0].source_file).suffix.lower()
    required_reads = 1 if source_suffix in IMAGE_SUFFIXES else 2
    if len(readable_ids) < required_reads:
        return None
    return {
        "passed": True,
        "plate_number": plate,
        "plate_side": side,
        "frame_ids": readable_ids,
    }


def call_doubao_review(
    candidate: dict,
    frames: list[Frame],
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
) -> dict:
    """对单个候选窗口做密集时序复核，并允许模型在硬闸门下读车牌。"""
    behavior_text = "、".join(SUPPORTED_BEHAVIORS)
    prompt = f"""
你在复核一段按时间顺序排列的原始视频帧。候选信息：
{json.dumps(remove_plate_fields(candidate), ensure_ascii=False)}

任务：
1. 判断候选是否属于以下轻度交通行为之一：{behavior_text}。不要判断事故责任，不处理严重伤亡画面。
2. 选择能证明行为、车辆轨迹和车牌的原始帧。
3. 只有前牌或后牌任意一面全部字符逐一清晰，且从该车牌帧到行为帧的车辆轨迹连续时，才输出 plate_number。视频必须在 plate_frame_ids 中选至少两张能独立读全号码的帧；单张图片任务可只选该图。
4. 任一字符模糊、镜头中断、车辆混淆或无法确认时，plate_number 必须为空，两个布尔值按事实填写。
5. 禁止生成式修复、补全或猜测车牌。

输出 JSON，不要 Markdown：
{{
  "event": {{
    "source_file": "...",
    "start_seconds": 0.0,
    "end_seconds": 0.0,
    "behavior": "...",
    "vehicle_type": "...",
    "vehicle_color": "...",
    "evidence_status": "待人工复核|证据不足",
    "evidence_note": "...",
    "evidence_frame_ids": ["..."]
  }},
  "plate_decision": {{
    "plate_number": "",
    "plate_side": "front|rear|",
    "plate_fully_clear": false,
    "trajectory_continuous": false,
    "plate_frame_ids": ["..."],
    "incident_frame_ids": ["..."],
    "ambiguity_reason": ""
  }}
}}
""".strip()
    content = [{"type": "text", "text": prompt}]
    for frame in frames:
        content.append(
            {
                "type": "text",
                "text": (
                    f"frame_id={frame.frame_id}; source={frame.source_file}; "
                    f"seconds={frame.timestamp:.3f}"
                ),
            }
        )
        content.append(
            {"type": "image_url", "image_url": {"url": image_data_url(frame.path)}}
        )
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": content}],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    response = post_json(
        f"{base_url.rstrip('/')}/chat/completions", api_key, payload, timeout
    )
    parsed = parse_json_content(extract_message_content(response))
    event = remove_plate_fields(parsed.get("event", {}))
    if not isinstance(event, dict):
        event = {}
    valid_frames = {frame.frame_id: frame for frame in frames}
    event["evidence_frame_ids"] = [
        frame_id
        for frame_id in event.get("evidence_frame_ids", [])
        if frame_id in valid_frames
    ][:5]
    initial = accept_model_plate(
        parsed.get("plate_decision"), valid_frames, require_consensus=False
    )
    accepted = None
    if initial:
        try:
            consensus = call_doubao_plate_consensus(
                initial, valid_frames, api_key, model, base_url, timeout
            )
        except (RuntimeError, ValueError, json.JSONDecodeError):
            consensus = None
        if consensus:
            initial["independent_consensus"] = consensus
            accepted = accept_model_plate(initial, valid_frames)
    if accepted:
        event["model_plate_decision"] = accepted
        event["evidence_frame_ids"] = list(
            dict.fromkeys(
                event["evidence_frame_ids"]
                + accepted["incident_frame_ids"]
                + accepted["plate_frame_ids"]
            )
        )[:5]
    elif isinstance(parsed.get("plate_decision"), dict):
        reason = str(parsed["plate_decision"].get("ambiguity_reason", "")).strip()
        if initial and not reason:
            reason = "独立车牌重读未达成逐字一致"
        if reason:
            event["plate_ambiguity_reason"] = PLATE_PATTERN.sub(
                "[车牌字符已移除]", reason
            )
    return event


def review_candidates(
    raw_candidates: list[dict],
    input_paths: list[Path],
    work_dir: Path,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
    dense_fps: float,
    padding: float,
    max_review_seconds: float,
    max_events: int,
) -> tuple[list[dict], list[Frame], dict[str, dict]]:
    path_by_name = {path.name: path for path in input_paths}
    reviewed: list[dict] = []
    dense_frames: list[Frame] = []
    plate_registry: dict[str, dict] = {}
    for index, candidate in enumerate(raw_candidates[:max_events], start=1):
        source = path_by_name.get(str(candidate.get("source_file", "")))
        if not source or source.suffix.lower() not in VIDEO_SUFFIXES:
            reviewed.append(remove_plate_fields(candidate))
            continue
        try:
            start = max(0.0, float(candidate.get("start_seconds", 0)) - padding)
            end = max(start + 0.5, float(candidate.get("end_seconds", start)) + padding)
        except (TypeError, ValueError):
            start, end = 0.0, max_review_seconds
        duration = min(max_review_seconds, end - start)
        frames = extract_video_window_frames(
            source,
            work_dir,
            dense_fps,
            start,
            duration,
            max(1, round(dense_fps * max_review_seconds)),
            f"review-{index:03d}",
        )
        if not frames:
            reviewed.append(remove_plate_fields(candidate))
            continue
        dense_frames.extend(frames)
        event = call_doubao_review(candidate, frames, api_key, model, base_url, timeout)
        decision = event.pop("model_plate_decision", None)
        if isinstance(decision, dict):
            plate_ref = f"plate-{index:03d}"
            plate_registry[plate_ref] = decision
            event["plate_ref"] = plate_ref
        reviewed.append(event)
    return reviewed, dense_frames, plate_registry


def normalize_with_deepseek(
    raw_candidates: list[dict],
    api_key: str,
    flash_model: str,
    pro_model: str,
    base_url: str,
    timeout: int,
) -> tuple[dict, str]:
    system_prompt = """
你只整理候选交通事件的文字 JSON。合并同一素材、时间重叠且车辆描述一致的重复候选。
不添加画面中没有的事实，不生成车牌字符，不将候选写成已认定违法。
输入中的 plate_ref 是本地已通过确定性车牌闸门的引用，只能原样保留或删除，不能改写或新造。
必须输出 JSON 对象，格式为：
{
  "review_required": false,
  "events": [
    {
      "event_id": "event-001",
      "source_file": "...",
      "start_seconds": 0.0,
      "end_seconds": 0.0,
      "behavior": "...",
      "vehicle_type": "...",
      "vehicle_color": "...",
      "evidence_status": "待人工复核|证据不足|模型判定车牌与轨迹通过（待人工复核）",
      "evidence_note": "...",
      "evidence_frame_ids": ["..."],
      "plate_ref": "plate-001|"
    }
  ]
}
只有输入候选相互冲突、且无法保守合并时，才设 review_required=true。
""".strip()
    user_prompt = "请整理以下 JSON 候选：\n" + json.dumps(
        raw_candidates, ensure_ascii=False
    )

    def invoke(model: str, extra_context: str = "") -> dict:
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt + extra_context},
            ],
            "temperature": 0,
            "response_format": {"type": "json_object"},
            "thinking": {"type": "disabled"},
        }
        response = post_json(
            f"{base_url.rstrip('/')}/chat/completions",
            api_key,
            payload,
            timeout,
        )
        return remove_plate_fields(
            parse_json_content(extract_message_content(response))
        )

    try:
        flash_result = invoke(flash_model)
        if not isinstance(flash_result.get("events"), list):
            raise ValueError("Flash 输出缺少 events 数组")
    except (RuntimeError, ValueError, json.JSONDecodeError) as flash_error:
        pro_result = invoke(
            pro_model,
            f"\nFlash 整理失败，请直接完成最终 JSON。错误类型：{type(flash_error).__name__}",
        )
        if not isinstance(pro_result.get("events"), list):
            raise ValueError("Pro 输出缺少 events 数组")
        return pro_result, pro_model

    if flash_result.get("review_required") is True:
        try:
            pro_result = invoke(
                pro_model,
                "\nFlash 认为存在冲突，请保守复核后输出最终 JSON。",
            )
            if isinstance(pro_result.get("events"), list):
                return pro_result, pro_model
        except (RuntimeError, ValueError, json.JSONDecodeError):
            pass
    return flash_result, flash_model


def attach_evidence_paths(
    analysis: dict,
    frames: list[Frame],
    plate_registry: dict[str, dict] | None = None,
) -> dict:
    registry = {frame.frame_id: frame for frame in frames}
    plate_registry = plate_registry or {}
    events = analysis.get("events", [])
    cleaned_events = []
    for index, event in enumerate(events, start=1):
        if not isinstance(event, dict):
            continue
        plate_ref = str(event.get("plate_ref", ""))
        event = remove_plate_fields(event)
        event["event_id"] = f"event-{index:03d}"
        frame_ids = [
            frame_id
            for frame_id in event.get("evidence_frame_ids", [])
            if frame_id in registry
        ]
        event["evidence_frame_ids"] = frame_ids[:5]
        mapped_images = [str(registry[frame_id].path) for frame_id in frame_ids[:5]]
        existing_images = [
            value
            for value in event.get("evidence_images", [])
            if isinstance(value, str) and Path(value).is_file()
        ]
        event["evidence_images"] = mapped_images or existing_images[:5]
        event["evidence_refs"] = [
            (
                f"{registry[frame_id].source_file}"
                f"@{registry[frame_id].timestamp:.3f}s"
            )
            for frame_id in frame_ids[:5]
        ]
        if not event.get("source_file") and frame_ids:
            event["source_file"] = registry[frame_ids[0]].source_file
        accepted_plate = plate_registry.get(plate_ref)
        if isinstance(accepted_plate, dict):
            event["plate_number"] = accepted_plate["plate_number"]
            event["plate_side"] = accepted_plate["plate_side"]
            event["plate_fully_clear"] = True
            event["trajectory_continuous"] = True
            event["evidence_status"] = "模型判定车牌与轨迹通过（待人工复核）"
        event.pop("plate_ref", None)
        cleaned_events.append(event)
    analysis["events"] = cleaned_events
    return analysis


def load_json(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"{path} 必须是 JSON 对象")
    return data


def main() -> None:
    parser = argparse.ArgumentParser(
        description="豆包 Seed 2.0 Lite -> DeepSeek V4 -> Excel 候选识别链路"
    )
    parser.add_argument("--input", nargs="+", required=True, help="视频或图片路径")
    parser.add_argument("--output", default="analysis.xlsx", help="输出 Excel 路径")
    parser.add_argument("--analysis-json-out", help="可选：保存最终规范化 JSON 供私有审计")
    parser.add_argument("--city", default="", help="用户确认的城市，仅写入运行上下文")
    parser.add_argument("--source-note", default="", help="素材来源/授权备注")
    parser.add_argument("--fps", type=float, default=2.0, help="完整时序扫描抽帧率，默认 2 fps")
    parser.add_argument("--dense-fps", type=float, default=4.0, help="候选窗口密集复核帧率")
    parser.add_argument("--candidate-padding", type=float, default=2.0)
    parser.add_argument("--max-review-seconds", type=float, default=8.0)
    parser.add_argument("--max-events", type=int, default=10)
    parser.add_argument("--max-frames", type=int, default=180, help="最多分析帧数")
    parser.add_argument(
        "--normalized-json",
        help="离线测试/人工接力：跳过两个模型，直接用规范化 JSON 生成 Excel",
    )
    parser.add_argument(
        "--ark-model",
        default=os.environ.get("DOUBAO_VISION_MODEL", "doubao-seed-2-0-lite-260215"),
    )
    parser.add_argument(
        "--ark-base-url",
        default=os.environ.get("ARK_BASE_URL", "https://ark.cn-beijing.volces.com/api/v3"),
    )
    parser.add_argument("--deepseek-flash-model", default="deepseek-v4-flash")
    parser.add_argument("--deepseek-pro-model", default="deepseek-v4-pro")
    parser.add_argument(
        "--deepseek-base-url",
        default=os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
    )
    parser.add_argument("--timeout", type=int, default=180)
    args = parser.parse_args()

    if min(args.fps, args.dense_fps, args.max_review_seconds) <= 0:
        parser.error("帧率和复核时长必须大于 0")
    if args.max_frames <= 0 or args.max_events <= 0:
        parser.error("--max-frames 和 --max-events 必须大于 0")

    input_paths = [Path(value).resolve() for value in args.input]
    output_path = Path(args.output).resolve()
    with tempfile.TemporaryDirectory(prefix="driving-justice-") as temp_dir:
        work_dir = Path(temp_dir)
        frames = prepare_frames(input_paths, work_dir, args.fps, args.max_frames)

        if args.normalized_json:
            offline = load_json(Path(args.normalized_json))
            plate_registry = {}
            for index, event in enumerate(offline.get("events", []), start=1):
                if not isinstance(event, dict):
                    continue
                accepted = accept_model_plate(
                    event.get("model_plate_decision"),
                    {frame.frame_id: frame for frame in frames},
                )
                if accepted:
                    ref = f"plate-{index:03d}"
                    plate_registry[ref] = accepted
                    event["plate_ref"] = ref
            analysis = remove_plate_fields(offline)
            route = "离线规范化 JSON -> Excel"
        else:
            ark_key = os.environ.get("ARK_API_KEY", "")
            deepseek_key = os.environ.get("DEEPSEEK_API_KEY", "")
            missing = [
                name
                for name, value in [
                    ("ARK_API_KEY", ark_key),
                    ("DEEPSEEK_API_KEY", deepseek_key),
                ]
                if not value
            ]
            if missing:
                raise RuntimeError(
                    "缺少 API 环境变量: " + ", ".join(missing)
                )
            raw_candidates = call_doubao_scan(
                frames,
                ark_key,
                args.ark_model,
                args.ark_base_url,
                args.timeout,
            )
            reviewed_candidates, dense_frames, plate_registry = review_candidates(
                raw_candidates,
                input_paths,
                work_dir,
                ark_key,
                args.ark_model,
                args.ark_base_url,
                args.timeout,
                args.dense_fps,
                args.candidate_padding,
                args.max_review_seconds,
                args.max_events,
            )
            frames.extend(dense_frames)
            analysis, deepseek_model = normalize_with_deepseek(
                reviewed_candidates,
                deepseek_key,
                args.deepseek_flash_model,
                args.deepseek_pro_model,
                args.deepseek_base_url,
                args.timeout,
            )
            route = f"{args.ark_model} -> {deepseek_model} -> Excel"

        analysis = attach_evidence_paths(analysis, frames, plate_registry)
        analysis["model_route"] = route
        analysis["source_summary"] = "；".join(path.name for path in input_paths)
        analysis["city"] = args.city
        analysis["source_note"] = args.source_note
        if args.analysis_json_out:
            analysis_json_path = Path(args.analysis_json_out).resolve()
            analysis_json_path.parent.mkdir(parents=True, exist_ok=True)
            analysis_json_path.write_text(
                json.dumps(analysis, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        write_workbook(analysis, output_path)

    print(f"已生成候选识别结果：{output_path}")
    print("注意：车牌仅在模型结构化判断通过确定性闸门时写入；正式举报前仍需用户确认。")


if __name__ == "__main__":
    try:
        main()
    except (RuntimeError, ValueError, FileNotFoundError, json.JSONDecodeError) as exc:
        sys.exit(f"执行失败：{exc}")
