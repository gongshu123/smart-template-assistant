from __future__ import annotations

import hashlib
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from PIL import Image
from psd_tools import PSDImage


FIELD_ORDER = ["image", "selling_point", "product_name", "price", "benefit", "discount", "sales"]
FIELD_LABELS = {
    "image": "产品图",
    "selling_point": "产品卖点（大标题）",
    "product_name": "产品名称（小标题）",
    "price": "价格",
    "benefit": "利益点",
    "discount": "立减",
    "sales": "销量",
}
FIELD_COLORS = {
    "image": "#2DBE70",
    "selling_point": "#A855F7",
    "product_name": "#2585E6",
    "price": "#EF4444",
    "benefit": "#14B8A6",
    "discount": "#F59E0B",
    "sales": "#EAB308",
}


@dataclass
class LayerInfo:
    layer_id: int | None
    parent_id: int | None
    name: str
    kind: str
    bbox: list[int] | None
    text: str | None
    smart_filename: str | None
    path_ids: list[int]
    path_names: list[str]
    depth: int
    visible: bool
    is_group: bool

    @property
    def center(self) -> tuple[float, float]:
        if not self.bbox:
            return (0.0, 0.0)
        return ((self.bbox[0] + self.bbox[2]) / 2, (self.bbox[1] + self.bbox[3]) / 2)

    @property
    def area(self) -> int:
        if not self.bbox:
            return 0
        return max(0, self.bbox[2] - self.bbox[0]) * max(0, self.bbox[3] - self.bbox[1])


@dataclass
class SlotMapping:
    slot_index: int
    root_layer_id: int | None
    root_name: str
    bbox: list[int] | None
    fields: dict[str, int | None]
    confidence: dict[str, str]


@dataclass
class TemplateAnalysis:
    template_path: str
    fingerprint: str
    width: int
    height: int
    total_layers: int
    hidden_layers: int
    layers: list[LayerInfo]
    slots: list[SlotMapping]

    def to_json(self) -> str:
        data = asdict(self)
        return json.dumps(data, ensure_ascii=False, indent=2)


def _bbox(layer: Any) -> list[int] | None:
    try:
        return [int(value) for value in layer.bbox]
    except Exception:
        return None


def _text(layer: Any) -> str | None:
    if getattr(layer, "kind", "") != "type":
        return None
    try:
        return str(layer.text).replace("\r", "\n")
    except Exception:
        return None


def _smart_filename(layer: Any) -> str | None:
    if getattr(layer, "kind", "") != "smartobject":
        return None
    try:
        return str(layer.smart_object.filename)
    except Exception:
        return None


def _walk(container: Any, parent_visible: bool = True, path_ids: tuple[int, ...] = (), path_names: tuple[str, ...] = (), parent_id: int | None = None) -> list[LayerInfo]:
    rows: list[LayerInfo] = []
    for layer in container:
        own_visible = bool(getattr(layer, "visible", True))
        effective_visible = parent_visible and own_visible
        layer_id = getattr(layer, "layer_id", None)
        current_ids = path_ids + ((int(layer_id),) if layer_id is not None else ())
        current_names = path_names + (str(getattr(layer, "name", "")),)
        row = LayerInfo(
            layer_id=int(layer_id) if layer_id is not None else None,
            parent_id=parent_id,
            name=str(getattr(layer, "name", "")),
            kind=str(getattr(layer, "kind", "")),
            bbox=_bbox(layer),
            text=_text(layer),
            smart_filename=_smart_filename(layer),
            path_ids=list(current_ids),
            path_names=list(current_names),
            depth=len(path_names),
            visible=effective_visible,
            is_group=bool(layer.is_group()),
        )
        rows.append(row)
        if layer.is_group():
            rows.extend(_walk(layer, effective_visible, current_ids, current_names, row.layer_id))
    return rows


def _descendant_bbox(layers: list[LayerInfo], root_id: int | None) -> list[int] | None:
    if root_id is None:
        return None
    boxes = [layer.bbox for layer in layers if layer.visible and not layer.is_group and layer.bbox and root_id in layer.path_ids[:-1]]
    if not boxes:
        return None
    return [min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes)]


def _fingerprint(path: Path) -> str:
    stat = path.stat()
    raw = f"{path.resolve()}|{stat.st_size}|{stat.st_mtime_ns}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()[:24]


def _common_prefix(paths: list[list[int]]) -> list[int]:
    if not paths:
        return []
    prefix = list(paths[0])
    for path in paths[1:]:
        length = 0
        for left, right in zip(prefix, path):
            if left != right:
                break
            length += 1
        prefix = prefix[:length]
    return prefix


def _product_anchors(layers: list[LayerInfo], width: int, height: int) -> list[LayerInfo]:
    visible_smart = [layer for layer in layers if layer.visible and layer.kind == "smartobject" and layer.bbox]
    code_pattern = re.compile(r"^[A-Z]{1,3}\d{5,}(?:\.[A-Za-z0-9]+)?$", re.I)
    semantic = []
    for layer in visible_smart:
        base = Path(layer.smart_filename or layer.name).name
        own_name = re.sub(r"\s+", "", layer.name)
        ancestor_names = [re.sub(r"\s+", "", name) for name in layer.path_names[:-1]]
        score = 0
        if code_pattern.match(base):
            score += 5
        if own_name in {"产品", "产品图", "商品", "商品图"}:
            score += 7
        if any(name in {"产品", "产品图", "商品", "商品图"} for name in ancestor_names):
            score += 5
        if layer.area > width * height * 0.015:
            score += 1
        if score >= 6:
            semantic.append((score, layer))

    candidates = [layer for _, layer in semantic]
    if len(candidates) >= 2:
        by_source: dict[str, list[LayerInfo]] = {}
        for layer in candidates:
            key = (layer.smart_filename or layer.name).lower()
            by_source.setdefault(key, []).append(layer)
        best = max(by_source.values(), key=lambda group: (len(group), sum(item.area for item in group)))
        if len(best) >= 2:
            return best
        return candidates

    by_source: dict[str, list[LayerInfo]] = {}
    for layer in visible_smart:
        key = (layer.smart_filename or layer.name).lower()
        by_source.setdefault(key, []).append(layer)
    repeated = [group for group in by_source.values() if len(group) >= 2]
    if repeated:
        return max(repeated, key=lambda group: (len(group), sum(item.area for item in group)))
    return candidates[:1]


def _is_fixed_text(value: str) -> bool:
    normalized = re.sub(r"\s+", "", value)
    fixed_tokens = ["下单立省", "拍下立省", "参考", "到手价", "为爱选购", "立即抢购", "悦己礼", "88VIP"]
    if normalized in {"￥", "¥", "元", "Ԫ", "起"}:
        return True
    return any(token in normalized for token in fixed_tokens)


def _classify_fields(slot_layers: list[LayerInfo], image: LayerInfo) -> tuple[dict[str, int | None], dict[str, str]]:
    fields: dict[str, int | None] = {field: None for field in FIELD_ORDER}
    confidence: dict[str, str] = {field: "未识别" for field in FIELD_ORDER}
    fields["image"] = image.layer_id
    confidence["image"] = "高"

    texts = [layer for layer in slot_layers if layer.visible and layer.kind == "type" and layer.text and layer.bbox]
    normalized = {layer.layer_id: re.sub(r"\s+", " ", layer.text or "").strip() for layer in texts}

    price_candidates = [layer for layer in texts if "价格" in layer.name and re.fullmatch(r"\s*\d+(?:\.\d+)?\s*", normalized[layer.layer_id])]
    if not price_candidates:
        price_candidates = [layer for layer in texts if re.fullmatch(r"\s*\d{2,}(?:\.\d+)?\s*", normalized[layer.layer_id])]
    if price_candidates:
        price = max(price_candidates, key=lambda layer: (layer.bbox[1], layer.area))
        fields["price"] = price.layer_id
        confidence["price"] = "高" if "价格" in price.name else "中"

    discount_candidates = [layer for layer in texts if re.fullmatch(r"\s*\d+(?:\.\d+)?\s*[元Ԫ]\s*", normalized[layer.layer_id])]
    if not discount_candidates:
        currency_parents = {
            layer.parent_id for layer in texts
            if re.fullmatch(r"\s*[元Ԫ]\s*", normalized[layer.layer_id])
        }
        discount_candidates = [
            layer for layer in texts
            if layer.parent_id in currency_parents
            and re.fullmatch(r"\s*\d+(?:\.\d+)?\s*", normalized[layer.layer_id])
            and any(
                other.parent_id == layer.parent_id and "立省" in normalized[other.layer_id]
                for other in texts
            )
        ]
    if discount_candidates:
        discount = min(discount_candidates, key=lambda layer: layer.bbox[1])
        fields["discount"] = discount.layer_id
        confidence["discount"] = "高"

    for field, keywords in {
        "benefit": ["利益点", "权益"],
        "sales": ["销量", "已售"],
    }.items():
        matches = [layer for layer in texts if any(keyword in layer.name or keyword in normalized[layer.layer_id] for keyword in keywords)]
        if matches:
            fields[field] = matches[0].layer_id
            confidence[field] = "高"

    excluded_ids = {value for value in fields.values() if value is not None}
    headline_candidates = []
    for layer in texts:
        if layer.layer_id in excluded_ids:
            continue
        value = normalized[layer.layer_id]
        if _is_fixed_text(value):
            continue
        if re.fullmatch(r"\d+(?:\.\d+)?", value):
            continue
        headline_candidates.append(layer)

    explicit_selling = [layer for layer in headline_candidates if "产品卖点" in layer.name or "大标题" in layer.name]
    explicit_name = [layer for layer in headline_candidates if "产品名称" in layer.name or "商品名称" in layer.name]
    if explicit_selling:
        fields["selling_point"] = explicit_selling[0].layer_id
        confidence["selling_point"] = "高"
    if explicit_name:
        fields["product_name"] = explicit_name[0].layer_id
        confidence["product_name"] = "高"

    remaining = [layer for layer in headline_candidates if layer.layer_id not in {fields["selling_point"], fields["product_name"]}]
    remaining.sort(key=lambda layer: (layer.bbox[1], -layer.area))
    if fields["selling_point"] is None and remaining:
        fields["selling_point"] = remaining[0].layer_id
        confidence["selling_point"] = "中"
        remaining = remaining[1:]
    if fields["product_name"] is None and remaining:
        selling = next((layer for layer in headline_candidates if layer.layer_id == fields["selling_point"]), None)
        below = [layer for layer in remaining if not selling or layer.bbox[1] >= selling.bbox[1]]
        selected = min(below or remaining, key=lambda layer: (abs(layer.bbox[1] - (selling.bbox[3] if selling else layer.bbox[1])), -layer.area))
        fields["product_name"] = selected.layer_id
        confidence["product_name"] = "中"

    return fields, confidence


def analyze_template(template_path: str | Path, preview_path: str | Path | None = None) -> TemplateAnalysis:
    path = Path(template_path)
    psd = PSDImage.open(path)
    layers = _walk(psd)
    visible_layers = [layer for layer in layers if layer.visible]
    anchors = _product_anchors(layers, psd.width, psd.height)
    anchors.sort(key=lambda layer: (round(layer.center[1] / 40), layer.center[0]))

    slots: list[SlotMapping] = []
    if anchors:
        layer_by_id = {layer.layer_id: layer for layer in layers if layer.layer_id is not None}
        anchor_ids = {anchor.layer_id for anchor in anchors if anchor.layer_id is not None}
        anchors_per_group: dict[int, int] = {}
        for layer in layers:
            if not layer.is_group or layer.layer_id is None:
                continue
            anchors_per_group[layer.layer_id] = sum(
                1 for anchor in anchors
                if anchor.layer_id in anchor_ids and layer.layer_id in anchor.path_ids[:-1]
            )
        for anchor in anchors:
            root_id = next(
                (group_id for group_id in anchor.path_ids[:-1] if anchors_per_group.get(group_id) == 1),
                anchor.parent_id,
            )
            root = layer_by_id.get(root_id)
            slot_layers = [layer for layer in visible_layers if root_id is None or root_id in layer.path_ids]
            fields, confidence = _classify_fields(slot_layers, anchor)
            slots.append(SlotMapping(
                slot_index=len(slots) + 1,
                root_layer_id=root_id,
                root_name=root.name if root else "商品位",
                bbox=_descendant_bbox(layers, root_id) or anchor.bbox,
                fields=fields,
                confidence=confidence,
            ))

    if preview_path:
        preview = psd.topil()
        if preview is None:
            preview = psd.composite()
        if preview is not None:
            preview_path = Path(preview_path)
            preview_path.parent.mkdir(parents=True, exist_ok=True)
            if preview.mode not in ("RGB", "RGBA"):
                preview = preview.convert("RGBA")
            preview.save(preview_path)

    return TemplateAnalysis(
        template_path=str(path.resolve()),
        fingerprint=_fingerprint(path),
        width=psd.width,
        height=psd.height,
        total_layers=len(layers),
        hidden_layers=len([layer for layer in layers if not layer.visible]),
        layers=layers,
        slots=slots,
    )


def save_analysis(analysis: TemplateAnalysis, path: str | Path) -> None:
    Path(path).write_text(analysis.to_json(), encoding="utf-8")
