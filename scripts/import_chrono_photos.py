#!/usr/bin/env python3
"""Incremental chronological photo importer for the life-log PWA / Google Apps Script backend."""

import argparse
import base64
from collections import defaultdict
from dataclasses import dataclass, field
import datetime
import hashlib
import io
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import time
import urllib.request

from PIL import Image, ImageOps

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8")

DEFAULT_DIR = Path(r"D:\電子相本\張文華的相本\年代照")
DEFAULT_TAG = "#我的年代照"
DEFAULT_STATE_PATH = Path(
    r"C:\Users\USER\.gemini\config\skills\life-log-chrono-importer\sync-state.json"
)
GAS_API_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbwaAPptJqC2a6yjgMas4VSr5D5YncKopX1h-n4joKtYnERwPvV6eCNOjYYKpErN_ao8/exec"
)
FILENAME_RE = re.compile(
    r"^(\d{3})(\d{2})(\d{2})?(?:_(\d+))?\.(jpg|jpeg|png|webp)$",
    re.IGNORECASE,
)
IGNORED_SYSTEM_FILES = {"thumbs.db", "desktop.ini", ".ds_store"}
MAX_DIM = 1920
SIZE_THRESHOLD_BYTES = 450 * 1024
JPEG_QUALITY = 85
MAX_PAYLOAD_BYTES = 5 * 1024 * 1024


@dataclass
class ParsedPhoto:
    """Metadata parsed from a chronological photo filename."""

    path: Path
    roc_year: int
    ad_year: int
    month: int
    day: int | None
    seq: int
    md5: str


@dataclass
class PhotoGroup:
    """Photos grouped by calendar date (or year-month) for a single life-log entry."""

    custom_id: str
    custom_date: str
    roc_year: int
    ad_year: int
    month: int
    day: int | None
    paths: list[Path]
    md5_map: dict[Path, str] = field(default_factory=dict)


@dataclass
class SyncAction:
    """Planned synchronization action for a PhotoGroup against remote logs."""

    action: str  # CREATE | APPEND | SKIP | WARN_SKIP
    group: PhotoGroup
    new_paths: list[Path]
    existing_entry: dict | None = None
    existing_urls: list[str] | None = None
    reason: str = ""


def now_iso_local() -> str:
    """Return current local ISO-8601 timestamp with timezone offset."""
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def normalize_tag(tag: str) -> str:
    """Ensure hashtag starts with '#' and is stripped of surrounding whitespace."""
    cleaned = (tag or "").strip()
    if not cleaned:
        return DEFAULT_TAG
    return cleaned if cleaned.startswith("#") else f"#{cleaned}"


def md5_file(path: Path) -> str:
    """Return hex MD5 digest of a local file."""
    h = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def get_group_md5(group: PhotoGroup, path: Path) -> str:
    """Return cached MD5 for a path in group, computing if missing."""
    cached = group.md5_map.get(path)
    if cached:
        return cached
    digest = md5_file(path)
    group.md5_map[path] = digest
    return digest


def load_sync_state(state_path: Path) -> dict:
    """Load local MD5 sync state JSON, returning empty structure if absent."""
    if not state_path.exists():
        return {"version": 1, "updated_at": "", "synced_md5": {}}
    data = json.loads(state_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        return {"version": 1, "updated_at": "", "synced_md5": {}}
    synced = data.get("synced_md5")
    if not isinstance(synced, dict):
        synced = {}
    return {
        "version": int(data.get("version", 1)),
        "updated_at": str(data.get("updated_at", "")),
        "synced_md5": synced,
    }


def save_sync_state(state_path: Path, state: dict) -> None:
    """Atomically persist local sync state JSON using temp file + os.replace."""
    state_path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix=".sync-state-",
        suffix=".tmp",
        dir=str(state_path.parent),
    )
    tmp_path = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(tmp_path, state_path)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        raise


def record_synced_paths(
    state: dict,
    group: PhotoGroup,
    paths: list[Path],
    state_path: Path,
) -> None:
    """Record uploaded/verified photo MD5 entries into local sync state and save."""
    timestamp = now_iso_local()
    synced_md5 = state.setdefault("synced_md5", {})
    for p in paths:
        digest = get_group_md5(group, p)
        synced_md5[digest] = {
            "filename": p.name,
            "custom_id": group.custom_id,
            "custom_date": group.custom_date,
            "synced_at": timestamp,
        }
    state["version"] = 1
    state["updated_at"] = timestamp
    save_sync_state(state_path, state)


def parse_photo_filename(path: Path) -> tuple[int, int, int, int | None, int] | None:
    """Parse ROC date components and sequence number from filename, or None if invalid."""
    m = FILENAME_RE.match(path.name)
    if not m:
        return None
    roc_y = int(m.group(1))
    ad_y = roc_y + 1911
    mo = int(m.group(2))
    da = int(m.group(3)) if m.group(3) else None
    seq = int(m.group(4)) if m.group(4) is not None else -1
    try:
        datetime.date(ad_y, mo, da or 1)
    except ValueError:
        return None
    return roc_y, ad_y, mo, da, seq


def scan_and_group_photos(
    candidate_paths: list[Path],
    strict: bool = False,
) -> tuple[list[PhotoGroup], list[Path], list[tuple[Path, Path]]]:
    """Parse, deduplicate by MD5, and group chronological photos into diary entries."""
    parsed_items: list[ParsedPhoto] = []
    unmatched: list[Path] = []

    for p in sorted(candidate_paths, key=lambda x: x.name.lower()):
        if not p.is_file():
            continue
        if p.name.lower() in IGNORED_SYSTEM_FILES:
            continue
        parsed = parse_photo_filename(p)
        if parsed is None:
            unmatched.append(p)
            continue
        roc_y, ad_y, mo, da, seq = parsed
        parsed_items.append(
            ParsedPhoto(
                path=p,
                roc_year=roc_y,
                ad_year=ad_y,
                month=mo,
                day=da,
                seq=seq,
                md5=md5_file(p),
            )
        )

    if strict and unmatched:
        bad_names = ", ".join(p.name for p in unmatched)
        raise ValueError(f"發現不符合命名規則的檔案（--strict 模式中止）：{bad_names}")

    # Deduplicate exact-content files, preferring explicit day over month-only, then earlier name.
    by_md5: dict[str, list[ParsedPhoto]] = defaultdict(list)
    for item in parsed_items:
        by_md5[item.md5].append(item)

    unique_items: list[ParsedPhoto] = []
    skipped_duplicates: list[tuple[Path, Path]] = []
    for _, dup_list in by_md5.items():
        if len(dup_list) == 1:
            unique_items.append(dup_list[0])
            continue
        ranked = sorted(
            dup_list,
            key=lambda it: (0 if it.day is not None else 1, it.path.name.lower()),
        )
        keeper = ranked[0]
        unique_items.append(keeper)
        for skipped in ranked[1:]:
            skipped_duplicates.append((skipped.path, keeper.path))

    grouped: dict[tuple[int, int, int | None], list[ParsedPhoto]] = defaultdict(list)
    for item in unique_items:
        grouped[(item.ad_year, item.month, item.day)].append(item)

    sorted_keys = sorted(grouped.keys(), key=lambda k: (k[0], k[1], k[2] or 1))
    groups: list[PhotoGroup] = []
    for ad_y, mo, da in sorted_keys:
        items = sorted(grouped[(ad_y, mo, da)], key=lambda it: (it.seq, it.path.name.lower()))
        roc_y = items[0].roc_year
        if da is not None:
            custom_id = f"chrono-{ad_y:04d}{mo:02d}{da:02d}"
            custom_date = f"{ad_y:04d}-{mo:02d}-{da:02d}"
        else:
            custom_id = f"chrono-{ad_y:04d}{mo:02d}01-ym"
            custom_date = f"{ad_y:04d}-{mo:02d}-01"
        groups.append(
            PhotoGroup(
                custom_id=custom_id,
                custom_date=custom_date,
                roc_year=roc_y,
                ad_year=ad_y,
                month=mo,
                day=da,
                paths=[it.path for it in items],
                md5_map={it.path: it.md5 for it in items},
            )
        )

    return groups, unmatched, skipped_duplicates


def build_initial_content(group: PhotoGroup) -> str:
    """Build standard diary text with ROC/AD date and original filenames."""
    names = "、".join(p.name for p in group.paths)
    if group.day is not None:
        header = (
            f"【我的年代照】民國{group.roc_year}年{group.month}月{group.day}日"
            f"（{group.ad_year:04d}/{group.month:02d}/{group.day:02d}）"
        )
    else:
        header = (
            f"【我的年代照】民國{group.roc_year}年{group.month}月"
            f"（{group.ad_year}年{group.month}月）"
        )
    if len(group.paths) > 1:
        header += f"（共 {len(group.paths)} 張）"
    return f"{header}\n原始檔名：{names}"


def is_filename_in_content(fname: str, content: str) -> bool:
    """Match filename cleanly without partial prefix/suffix collisions."""
    if not content:
        return False
    pattern = rf"(?<![A-Za-z0-9_]){re.escape(fname)}(?![A-Za-z0-9_])"
    return re.search(pattern, content, re.IGNORECASE) is not None


def parse_existing_urls(image_field: str) -> list[str]:
    """Split newline/whitespace-separated Drive URLs from a sheet row."""
    if not image_field:
        return []
    return [u.strip() for u in re.split(r"[\r\n\s]+", str(image_field)) if u.strip()]


def merge_content_for_append(
    existing_content: str,
    new_names: list[str],
    total_photo_count: int,
) -> str:
    """Update photo count and append filenames if marker exists, keeping handwritten stories intact."""
    text = existing_content or ""
    lines = text.splitlines()
    joined_new = "、".join(new_names)

    if lines:
        first_line = lines[0]
        if re.search(r"（共\s*\d+\s*張）", first_line):
            lines[0] = re.sub(r"（共\s*\d+\s*張）", f"（共 {total_photo_count} 張）", first_line)
        elif first_line.startswith("【我的年代照】") and total_photo_count > 1:
            lines[0] = f"{first_line}（共 {total_photo_count} 張）"

    for idx, line in enumerate(lines):
        if line.startswith("原始檔名：") or line.startswith("原始檔名:"):
            stripped = line.rstrip()
            if stripped.endswith(("：", ":", "、")):
                lines[idx] = f"{stripped}{joined_new}"
            else:
                lines[idx] = f"{stripped}、{joined_new}"
            break

    return "\n".join(lines)


def merge_tags(existing_tags: str, target_tag: str) -> str:
    """Preserve all existing tags and ensure target_tag is included once."""
    existing_str = (existing_tags or "").strip()
    if not existing_str:
        return target_tag
    tokens = [t for t in re.split(r"[\s,]+", existing_str) if t]
    if target_tag.lower() not in {t.lower() for t in tokens}:
        tokens.append(target_tag)
    return " ".join(tokens)


def plan_sync_actions(
    groups: list[PhotoGroup],
    existing_by_id: dict[str, dict],
    tag: str,
    is_files_mode: bool = False,
    sync_state: dict | None = None,
    force: bool = False,
) -> list[SyncAction]:
    """Compare local photo groups against local MD5 sync state and cloud records."""
    del tag
    synced_md5 = (sync_state or {}).get("synced_md5") or {}
    synced_filenames = {
        str(v.get("filename", "")).lower()
        for v in synced_md5.values()
        if isinstance(v, dict) and v.get("filename")
    }
    has_local_state = (not force) and len(synced_md5) > 0
    actions: list[SyncAction] = []

    for group in groups:
        if force:
            unsynced_paths = list(group.paths)
            modified_paths: list[Path] = []
        else:
            unsynced_paths = []
            modified_paths = []
            for p in group.paths:
                digest = get_group_md5(group, p)
                if digest in synced_md5:
                    continue
                if p.name.lower() in synced_filenames:
                    modified_paths.append(p)
                else:
                    unsynced_paths.append(p)

        if not unsynced_paths:
            if modified_paths:
                mod_names = "、".join(p.name for p in modified_paths)
                actions.append(
                    SyncAction(
                        action="WARN_SKIP",
                        group=group,
                        new_paths=[],
                        reason=(
                            f"檔名已曾同步但內容變更（{mod_names}），"
                            f"若確定要補傳請用 --files <路徑> --force"
                        ),
                    )
                )
            else:
                actions.append(
                    SyncAction(
                        action="SKIP",
                        group=group,
                        new_paths=[],
                        reason="本機狀態檔已記錄為已同步",
                    )
                )
            continue

        existing = existing_by_id.get(group.custom_id)
        if existing is None:
            actions.append(
                SyncAction(
                    action="CREATE",
                    group=group,
                    new_paths=unsynced_paths,
                    reason="雲端尚無此日期年代照",
                )
            )
            continue

        existing_content = str(existing.get("content") or "")
        existing_urls = parse_existing_urls(existing.get("image") or "")
        has_filename_marker = (
            "原始檔名：" in existing_content
            or "原始檔名:" in existing_content
            or any(is_filename_in_content(p.name, existing_content) for p in group.paths)
        )

        if force:
            candidates = unsynced_paths
        else:
            candidates = [
                p for p in unsynced_paths if not is_filename_in_content(p.name, existing_content)
            ]

        if not candidates:
            if modified_paths:
                mod_names = "、".join(p.name for p in modified_paths)
                actions.append(
                    SyncAction(
                        action="WARN_SKIP",
                        group=group,
                        new_paths=[],
                        existing_entry=existing,
                        existing_urls=existing_urls,
                        reason=(
                            f"檔名已曾同步但內容變更（{mod_names}），"
                            f"若確定要補傳請用 --files <路徑> --force"
                        ),
                    )
                )
            else:
                actions.append(
                    SyncAction(
                        action="SKIP",
                        group=group,
                        new_paths=[],
                        existing_entry=existing,
                        existing_urls=existing_urls,
                        reason="雲端內文已含此檔名",
                    )
                )
            continue

        if has_local_state or has_filename_marker or is_files_mode or force:
            note = f"補上傳 {len(candidates)} 張新照片至既有相簿"
            if modified_paths:
                note += f"（另有內容變更檔略過：{', '.join(p.name for p in modified_paths)}）"
            actions.append(
                SyncAction(
                    action="APPEND",
                    group=group,
                    new_paths=candidates,
                    existing_entry=existing,
                    existing_urls=existing_urls,
                    reason=note,
                )
            )
        elif len(group.paths) <= len(existing_urls):
            actions.append(
                SyncAction(
                    action="SKIP",
                    group=group,
                    new_paths=[],
                    existing_entry=existing,
                    existing_urls=existing_urls,
                    reason="已自訂內文且雲端照片數未少於資料夾",
                )
            )
        else:
            actions.append(
                SyncAction(
                    action="WARN_SKIP",
                    group=group,
                    new_paths=[],
                    existing_entry=existing,
                    existing_urls=existing_urls,
                    reason=(
                        f"無本機狀態檔且雲端已移除「原始檔名：」"
                        f"（本地 {len(group.paths)} 張 > 雲端 {len(existing_urls)} 張），"
                        f"為避免重複上傳舊照已略過；請改用 --files 指定新增檔案"
                    ),
                )
            )

    return actions


def optimize_image(path: Path) -> dict:
    """Match PWA compressImage behavior while preserving small standard JPEGs."""
    raw_bytes = path.read_bytes()
    raw_size = len(raw_bytes)

    with Image.open(path) as im:
        fmt = (im.format or "").upper()
        w, h = im.size
        exif = im.getexif()
        orientation = exif.get(274, 1) if exif else 1

        needs_process = (
            fmt != "JPEG"
            or max(w, h) > MAX_DIM
            or raw_size > SIZE_THRESHOLD_BYTES
            or orientation != 1
        )

        if not needs_process:
            b64 = base64.b64encode(raw_bytes).decode("ascii")
            return {
                "name": path.name,
                "type": "image/jpeg",
                "data": f"data:image/jpeg;base64,{b64}",
                "raw_bytes": raw_size,
                "out_bytes": raw_size,
                "processed": False,
                "orig_fmt": fmt,
                "orig_size": f"{w}x{h}",
                "out_size": f"{w}x{h}",
            }

        img = ImageOps.exif_transpose(im)
        if img.mode == "P":
            img = img.convert("RGBA" if "transparency" in img.info else "RGB")
        elif img.mode == "LA":
            img = img.convert("RGBA")

        if "A" in img.mode:
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")

        pw, ph = img.size
        if max(pw, ph) > MAX_DIM:
            img.thumbnail((MAX_DIM, MAX_DIM), Image.Resampling.LANCZOS)
        ow, oh = img.size

        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        out_bytes = buf.getvalue()

        b64 = base64.b64encode(out_bytes).decode("ascii")
        return {
            "name": f"{path.stem}.jpg",
            "type": "image/jpeg",
            "data": f"data:image/jpeg;base64,{b64}",
            "raw_bytes": raw_size,
            "out_bytes": len(out_bytes),
            "processed": True,
            "orig_fmt": fmt,
            "orig_size": f"{w}x{h}",
            "out_size": f"{ow}x{oh}",
        }


def post_gas(action: str, data: dict, timeout: int = 120) -> dict:
    """Send POST JSON payload to GAS Web App and return decoded JSON response."""
    body = json.dumps({"action": action, "data": data}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        GAS_API_URL,
        data=body,
        headers={"Content-Type": "text/plain;charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def fetch_existing_by_id(tag: str) -> dict[str, dict]:
    """Fetch existing diaries matching tag from GAS and index them by entry id."""
    resp = post_gas("searchLogs", {"keyword": tag, "searchDate": ""}, timeout=60)
    if not isinstance(resp, dict) or resp.get("success") is False:
        raise RuntimeError(f"查詢雲端既有日記失敗：{resp}")
    rows = resp.get("data") or []
    return {str(r["id"]): r for r in rows if r and r.get("id")}


def post_with_retry(action: str, payload: dict, validate_fn) -> dict:
    """Post to GAS with exponential backoff retries (up to 4 attempts total)."""
    last_err = None
    for attempt in range(1, 5):
        try:
            resp = post_gas(action, payload, timeout=120)
            validate_fn(resp)
            return resp
        except Exception as exc:
            last_err = exc
            if attempt < 4:
                wait_s = 2 ** attempt
                print(f"  [重試 {attempt}/3] 發生錯誤 ({exc})，{wait_s} 秒後重試...")
                time.sleep(wait_s)
    raise RuntimeError(f"重試 3 次後仍失敗：{last_err}")


def print_plan_summary(
    actions: list[SyncAction],
    unmatched: list[Path],
    skipped_dups: list[tuple[Path, Path]],
) -> None:
    """Print detailed action plan and counts."""
    for skipped_p, kept_p in skipped_dups:
        print(f"[去重略過] {skipped_p.name} 與 {kept_p.name} 內容完全相同，保留較精確/較前的 {kept_p.name}")
    for bad_p in unmatched:
        print(f"[警告略過] 不符合民國檔名規則：{bad_p.name}")

    for idx, act in enumerate(actions, 1):
        g = act.group
        if act.action == "CREATE":
            names = "、".join(p.name for p in act.new_paths)
            print(
                f"[{idx:02d}/{len(actions):02d}] CREATE    {g.custom_date} | {g.custom_id:<19} | "
                f"+{len(act.new_paths):2d} 張 | {names}"
            )
        elif act.action == "APPEND":
            names = "、".join(p.name for p in act.new_paths)
            existing_n = len(act.existing_urls or [])
            print(
                f"[{idx:02d}/{len(actions):02d}] APPEND    {g.custom_date} | {g.custom_id:<19} | "
                f"既有 {existing_n} 張 + {len(act.new_paths)} 張 | {names}"
            )
        elif act.action == "WARN_SKIP":
            print(
                f"[{idx:02d}/{len(actions):02d}] WARN_SKIP {g.custom_date} | {g.custom_id:<19} | {act.reason}"
            )
        else:
            print(
                f"[{idx:02d}/{len(actions):02d}] SKIP      {g.custom_date} | {g.custom_id:<19} | "
                f"{len(g.paths):2d} 張 ({act.reason})"
            )

    creates = [a for a in actions if a.action == "CREATE"]
    appends = [a for a in actions if a.action == "APPEND"]
    skips = [a for a in actions if a.action == "SKIP"]
    warns = [a for a in actions if a.action == "WARN_SKIP"]

    print("\n=== 掃描與比對摘要 ===")
    print(f"預計新增 (CREATE)      : {len(creates)} 篇（{sum(len(a.new_paths) for a in creates)} 張照片）")
    print(f"預計補圖 (APPEND)      : {len(appends)} 篇（{sum(len(a.new_paths) for a in appends)} 張照片）")
    print(f"已存在略過 (SKIP)      : {len(skips)} 篇（{sum(len(a.group.paths) for a in skips)} 張照片）")
    print(f"需人工確認略過 (WARN)  : {len(warns)} 篇")
    print(f"MD5 重複自動略過       : {len(skipped_dups)} 檔")
    print(f"異常命名警告 (unmatched): {len(unmatched)} 檔")


def execute_sync(
    actions: list[SyncAction],
    tag: str,
    sync_state: dict,
    state_path: Path,
    no_sort: bool = False,
) -> tuple[int, int, int]:
    """Execute CREATE and APPEND actions against GAS API, persisting state after each success."""
    created_count = 0
    appended_count = 0
    uploaded_photos = 0
    consecutive_failures = 0

    work_items = [a for a in actions if a.action in ("CREATE", "APPEND")]
    if not work_items:
        print("\n沒有需要新增或補圖的項目。")
        return 0, 0, 0

    print(f"\n開始執行同步：共 {len(work_items)} 筆變更...")
    for idx, act in enumerate(work_items, 1):
        g = act.group
        files_meta = [optimize_image(p) for p in act.new_paths]
        api_files = [{"name": m["name"], "type": m["type"], "data": m["data"]} for m in files_meta]
        expected_n = len(api_files)

        try:
            if act.action == "CREATE":
                payload = {
                    "customId": g.custom_id,
                    "customDate": g.custom_date,
                    "customTime": "00:00",
                    "content": build_initial_content(g),
                    "tags": tag,
                    "skipAI": True,
                    "skipSort": True,
                    "files": api_files,
                }
                payload_bytes = len(json.dumps(payload, ensure_ascii=False).encode("utf-8"))
                if payload_bytes >= MAX_PAYLOAD_BYTES:
                    raise RuntimeError(f"{g.custom_id} payload 過大 ({payload_bytes} bytes)")

                def _validate_create(resp: dict) -> None:
                    if not isinstance(resp, dict) or not resp.get("success"):
                        raise RuntimeError(f"saveDiaryV8 回傳失敗：{resp}")
                    if not resp.get("skipped"):
                        got = (resp.get("data") or {}).get("uploadedCount")
                        if got != expected_n:
                            raise RuntimeError(f"上傳張數不符：預期 {expected_n}，實得 {got}")

                resp = post_with_retry("saveDiaryV8", payload, _validate_create)
                consecutive_failures = 0
                record_synced_paths(sync_state, g, act.new_paths, state_path)
                if resp.get("skipped"):
                    print(f"[{idx}/{len(work_items)}] SKIPPED (雲端已存在，已補寫狀態) {g.custom_id}")
                else:
                    created_count += 1
                    uploaded_photos += expected_n
                    print(f"[{idx}/{len(work_items)}] CREATED {g.custom_id} (+{expected_n} 張)")
            else:
                existing = act.existing_entry or {}
                existing_urls = act.existing_urls or []
                total_after = len(existing_urls) + expected_n
                updated_content = merge_content_for_append(
                    str(existing.get("content") or ""),
                    [p.name for p in act.new_paths],
                    total_after,
                )
                updated_tags = merge_tags(str(existing.get("tags") or ""), tag)
                payload = {
                    "id": g.custom_id,
                    "content": updated_content,
                    "tags": updated_tags,
                    "files": api_files,
                    "keepImageUrls": existing_urls,
                    "deleteFileIds": [],
                    "deleteImage": False,
                }
                payload_bytes = len(json.dumps(payload, ensure_ascii=False).encode("utf-8"))
                if payload_bytes >= MAX_PAYLOAD_BYTES:
                    raise RuntimeError(f"{g.custom_id} payload 過大 ({payload_bytes} bytes)")

                def _validate_update(resp: dict) -> None:
                    if not isinstance(resp, dict) or not resp.get("success"):
                        raise RuntimeError(f"updateLog 回傳失敗：{resp}")

                post_with_retry("updateLog", payload, _validate_update)
                consecutive_failures = 0
                record_synced_paths(sync_state, g, act.new_paths, state_path)
                appended_count += 1
                uploaded_photos += expected_n
                print(
                    f"[{idx}/{len(work_items)}] APPENDED {g.custom_id} "
                    f"(既有 {len(existing_urls)} + 新增 {expected_n} -> 共 {total_after} 張)"
                )
        except Exception as exc:
            consecutive_failures += 1
            print(f"[{idx}/{len(work_items)}] FAILED {g.custom_id}: {exc}", file=sys.stderr)
            if consecutive_failures >= 3:
                raise SystemExit("連續失敗達 3 次，為避免異常已中止執行。")

        time.sleep(0.8)

    if created_count > 0 and not no_sort:
        print("\n偵測到新增年代照日記，正在執行全表日期排序 (sortLogs)...")
        sort_resp = post_with_retry(
            "sortLogs",
            {},
            lambda r: None if (isinstance(r, dict) and r.get("success")) else (_ for _ in ()).throw(
                RuntimeError(f"sortLogs 失敗：{r}")
            ),
        )
        print(f"sortLogs 完成：{sort_resp.get('message')}")

    print(
        f"\n=== 同步完成 ===\n"
        f"新增日記：{created_count} 篇｜補圖更新：{appended_count} 篇｜共上傳照片：{uploaded_photos} 張"
    )
    return created_count, appended_count, uploaded_photos


def mark_groups_synced(
    groups: list[PhotoGroup],
    sync_state: dict,
    state_path: Path,
) -> int:
    """Register all deduplicated photos into local sync state without uploading."""
    timestamp = now_iso_local()
    synced_md5 = sync_state.setdefault("synced_md5", {})
    marked_count = 0
    for g in groups:
        for p in g.paths:
            digest = get_group_md5(g, p)
            synced_md5[digest] = {
                "filename": p.name,
                "custom_id": g.custom_id,
                "custom_date": g.custom_date,
                "synced_at": timestamp,
            }
            marked_count += 1
    sync_state["version"] = 1
    sync_state["updated_at"] = timestamp
    save_sync_state(state_path, sync_state)
    return marked_count


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="將民國年代照增量匯入 life-log PWA（支援同日合併、MD5 狀態追蹤、補圖保留手寫內文與自動排序）。"
    )
    source_group = parser.add_mutually_exclusive_group()
    source_group.add_argument(
        "--dir",
        type=Path,
        default=None,
        help=f"年代照資料夾路徑（預設：{DEFAULT_DIR}）",
    )
    source_group.add_argument(
        "--files",
        type=Path,
        nargs="+",
        default=None,
        help="指定單張或多張年代照檔案路徑（與 --dir 互斥）",
    )
    parser.add_argument(
        "--tag",
        type=str,
        default=DEFAULT_TAG,
        help=f"寫入的標籤（預設：{DEFAULT_TAG}）",
    )
    parser.add_argument(
        "--state",
        type=Path,
        default=DEFAULT_STATE_PATH,
        help=f"本機 MD5 同步狀態檔路徑（預設：{DEFAULT_STATE_PATH}）",
    )
    parser.add_argument(
        "--mark-synced",
        action="store_true",
        help="將掃描到的不重複照片直接登記為已同步並寫入狀態檔（不上傳至雲端）",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="忽略本機狀態檔檢查，強制重新比對或補傳（建議搭配 --files 使用）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="僅掃描與比對本機／雲端狀態，不實際上傳或排序",
    )
    parser.add_argument(
        "--no-sort",
        action="store_true",
        help="跳過匯入完成後的 sortLogs 日期排序",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="若資料夾內出現不符合命名規則的檔案立即中止",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    tag = normalize_tag(args.tag)
    state_path = args.state.resolve()

    is_files_mode = args.files is not None
    if is_files_mode:
        candidate_paths = [p.resolve() for p in args.files]
        missing = [p for p in candidate_paths if not p.is_file()]
        if missing:
            print(f"找不到指定檔案：{', '.join(str(m) for m in missing)}", file=sys.stderr)
            return 1
        print(f"來源模式：指定檔案（共 {len(candidate_paths)} 個）")
    else:
        target_dir = (args.dir or DEFAULT_DIR).resolve()
        if not target_dir.is_dir():
            print(f"找不到資料夾：{target_dir}", file=sys.stderr)
            return 1
        candidate_paths = sorted([p for p in target_dir.iterdir() if p.is_file()], key=lambda p: p.name)
        print(f"來源資料夾：{target_dir}（掃描到 {len(candidate_paths)} 個檔案）")

    try:
        groups, unmatched, skipped_dups = scan_and_group_photos(candidate_paths, strict=args.strict)
    except ValueError as exc:
        print(f"錯誤：{exc}", file=sys.stderr)
        return 1

    print(f"解析完成：{len(groups)} 個日期群組（含 {sum(len(g.paths) for g in groups)} 張不重複照片）")
    sync_state = load_sync_state(state_path)

    if args.mark_synced:
        for skipped_p, kept_p in skipped_dups:
            print(f"[去重略過] {skipped_p.name} 與 {kept_p.name} 內容完全相同，保留較精確/較前的 {kept_p.name}")
        for bad_p in unmatched:
            print(f"[警告略過] 不符合民國檔名規則：{bad_p.name}")
        marked = mark_groups_synced(groups, sync_state, state_path)
        print(
            f"已標記完成：{len(groups)} 篇 / {marked} 張不重複照片已登記為已同步\n"
            f"狀態檔路徑：{state_path}（累計 {len(sync_state['synced_md5'])} 筆 MD5 紀錄）"
        )
        return 0

    print(
        f"本機狀態檔：{state_path}（"
        f"{'已載入 ' + str(len(sync_state['synced_md5'])) + ' 筆 MD5' if sync_state['synced_md5'] else '尚未初始化'}）"
    )
    existing_by_id = fetch_existing_by_id(tag)
    print(f"雲端既有 {tag} 日記：{len(existing_by_id)} 篇\n")

    actions = plan_sync_actions(
        groups,
        existing_by_id,
        tag,
        is_files_mode=is_files_mode,
        sync_state=sync_state,
        force=args.force,
    )
    print_plan_summary(actions, unmatched, skipped_dups)

    if args.dry_run:
        print("\n（--dry-run 模式：未上傳任何資料）")
        return 0

    execute_sync(
        actions,
        tag=tag,
        sync_state=sync_state,
        state_path=state_path,
        no_sort=args.no_sort,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
