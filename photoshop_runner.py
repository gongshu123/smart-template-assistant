from __future__ import annotations

import json
import os
import re
import subprocess
from pathlib import Path

from analyzer import LayerInfo, TemplateAnalysis
from excel_parser import WorkbookData


def find_photoshop() -> Path | None:
    roots = [Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Adobe"]
    candidates: list[Path] = []
    for root in roots:
        if not root.exists():
            continue
        candidates.extend(root.glob("Adobe Photoshop */Photoshop.exe"))
    if not candidates:
        return None
    return sorted(candidates, key=lambda path: path.parent.name, reverse=True)[0]


def _safe_name(name: str) -> str:
    cleaned = re.sub(r"[<>:\\|?*\"/]", "_", name).strip().rstrip(".")
    return cleaned or "套版输出"


def _unique_output(folder: Path, name: str, reserved: set[Path]) -> Path:
    candidate = folder / f"{_safe_name(name)}.psd"
    index = 2
    while candidate.exists() or candidate in reserved:
        candidate = folder / f"{_safe_name(name)}_{index}.psd"
        index += 1
    reserved.add(candidate)
    return candidate


def _layer_by_id(analysis: TemplateAnalysis) -> dict[int, LayerInfo]:
    return {layer.layer_id: layer for layer in analysis.layers if layer.layer_id is not None}


def _slot_job(analysis: TemplateAnalysis, slot_index: int) -> dict:
    slot = analysis.slots[slot_index]
    by_id = _layer_by_id(analysis)
    fields = dict(slot.fields)

    price_aux: list[int] = []
    price_id = fields.get("price")
    if price_id and price_id in by_id:
        price_layer = by_id[price_id]
        for layer in analysis.layers:
            if layer.visible and layer.kind == "type" and layer.parent_id == price_layer.parent_id:
                text = (layer.text or "").strip()
                if text in {"￥", "¥"} and layer.layer_id is not None:
                    price_aux.append(layer.layer_id)

    discount_group_id = None
    discount_aux: list[int] = []
    discount_id = fields.get("discount")
    if discount_id and discount_id in by_id:
        discount_layer = by_id[discount_id]
        discount_group_id = discount_layer.parent_id
        for layer in analysis.layers:
            if layer.visible and layer.kind == "type" and layer.parent_id == discount_layer.parent_id:
                text = (layer.text or "").strip()
                if text in {"元", "Ԫ"} and layer.layer_id is not None:
                    discount_aux.append(layer.layer_id)

    return {
        "fields": fields,
        "priceAux": price_aux,
        "discountGroupId": discount_group_id,
        "discountAux": discount_aux,
    }


def build_jobs(analysis: TemplateAnalysis, workbook: WorkbookData, asset_folder: str | Path, output_folder: str | Path) -> list[dict]:
    assets = Path(asset_folder).resolve()
    outputs = Path(output_folder).resolve()
    outputs.mkdir(parents=True, exist_ok=True)
    reserved: set[Path] = set()
    slot_defs = [_slot_job(analysis, index) for index in range(len(analysis.slots))]

    jobs = []
    for page in workbook.pages:
        output_path = _unique_output(outputs, page.name, reserved)
        products = []
        for index, product in enumerate(page.products[:len(slot_defs)]):
            values = vars(product).copy()
            if values.get("image"):
                values["image"] = str((assets / values["image"]).resolve()).replace("\\", "/")
            products.append({"slot": slot_defs[index], "values": values})
        jobs.append({
            "name": page.name,
            "template": str(Path(analysis.template_path).resolve()).replace("\\", "/"),
            "output": str(output_path).replace("\\", "/"),
            "products": products,
        })
    return jobs


JSX_ENGINE = r'''#target photoshop
(function () {
    var JOBS = __JOBS__;
    var LOG_PATH = __LOG_PATH__;
    var logFile = new File(LOG_PATH);
    logFile.encoding = "UTF8";
    logFile.parent.create();
    logFile.open("w");

    function log(message) {
        logFile.writeln(new Date().toString() + " " + message);
        logFile.close();
        logFile.open("a");
    }

    function isBlank(value) {
        if (value === null || value === undefined) return true;
        var text = String(value).replace(/^\s+|\s+$/g, "");
        return text === "" || /^0(?:\.0+)?$/.test(text);
    }

    function findLayerById(container, id) {
        if (!id) return null;
        for (var i = 0; i < container.layers.length; i++) {
            var layer = container.layers[i];
            try { if (layer.id === id) return layer; } catch (e) {}
            if (layer.typename === "LayerSet") {
                var nested = findLayerById(layer, id);
                if (nested) return nested;
            }
        }
        return null;
    }

    function requireLayer(document, id, label) {
        var layer = findLayerById(document, id);
        if (!layer) throw new Error("找不到" + label + "图层，ID=" + id);
        return layer;
    }

    function setSimpleText(document, id, value) {
        if (!id) return;
        var layer = requireLayer(document, id, "文字");
        if (isBlank(value)) {
            layer.visible = false;
            return;
        }
        layer.visible = true;
        layer.textItem.contents = String(value);
    }

    function setTextPreserveTrailingStyle(document, layer, bodyText, suffixText) {
        document.activeLayer = layer;
        var layerReference = new ActionReference();
        layerReference.putIdentifier(charIDToTypeID("Lyr "), layer.id);
        var layerDescriptor = executeActionGet(layerReference);
        var textKeyId = stringIDToTypeID("textKey");
        var textDescriptor = layerDescriptor.getObjectValue(textKeyId);
        var styleRangeId = stringIDToTypeID("textStyleRange");
        var fromId = stringIDToTypeID("from");
        var toId = stringIDToTypeID("to");
        var oldStyles = textDescriptor.getList(styleRangeId);
        var numberStyle = oldStyles.getObjectValue(0);
        var suffixStyle = oldStyles.getObjectValue(oldStyles.count > 1 ? oldStyles.count - 1 : 0);
        var body = String(bodyText);
        var fullText = body + String(suffixText);
        numberStyle.putInteger(fromId, 0);
        numberStyle.putInteger(toId, body.length);
        suffixStyle.putInteger(fromId, body.length);
        suffixStyle.putInteger(toId, fullText.length + 1);
        var newStyles = new ActionList();
        newStyles.putObject(styleRangeId, numberStyle);
        newStyles.putObject(styleRangeId, suffixStyle);
        textDescriptor.putString(textKeyId, fullText);
        textDescriptor.putList(styleRangeId, newStyles);
        var paragraphRangeId = stringIDToTypeID("paragraphStyleRange");
        if (textDescriptor.hasKey(paragraphRangeId)) {
            var oldParagraphs = textDescriptor.getList(paragraphRangeId);
            var newParagraphs = new ActionList();
            for (var p = 0; p < oldParagraphs.count; p++) {
                var paragraph = oldParagraphs.getObjectValue(p);
                if (p === 0) paragraph.putInteger(fromId, 0);
                if (p === oldParagraphs.count - 1) paragraph.putInteger(toId, fullText.length + 1);
                newParagraphs.putObject(paragraphRangeId, paragraph);
            }
            textDescriptor.putList(paragraphRangeId, newParagraphs);
        }
        var setDescriptor = new ActionDescriptor();
        setDescriptor.putReference(charIDToTypeID("null"), layerReference);
        setDescriptor.putObject(charIDToTypeID("T   "), charIDToTypeID("TxLr"), textDescriptor);
        executeAction(charIDToTypeID("setd"), setDescriptor, DialogModes.NO);
    }

    function setPrice(document, slot, value) {
        var priceId = slot.fields.price;
        if (!priceId) return;
        var price = requireLayer(document, priceId, "价格");
        var blank = isBlank(value);
        price.visible = !blank;
        if (!blank) price.textItem.contents = String(value);
        for (var i = 0; i < slot.priceAux.length; i++) {
            var auxiliary = findLayerById(document, slot.priceAux[i]);
            if (auxiliary) auxiliary.visible = !blank;
        }
    }

    function setDiscount(document, slot, value) {
        var discountId = slot.fields.discount;
        if (!discountId) return;
        var blank = isBlank(value);
        var group = slot.discountGroupId ? findLayerById(document, slot.discountGroupId) : null;
        if (group) group.visible = !blank;
        var layer = requireLayer(document, discountId, "立减");
        if (blank) {
            layer.visible = false;
            return;
        }
        layer.visible = true;
        var text = String(value).replace(/[元Ԫ]\s*$/, "");
        if (slot.discountAux && slot.discountAux.length) {
            layer.textItem.contents = text;
            for (var i = 0; i < slot.discountAux.length; i++) {
                var suffixLayer = findLayerById(document, slot.discountAux[i]);
                if (suffixLayer) suffixLayer.visible = true;
            }
        } else {
            setTextPreserveTrailingStyle(document, layer, text, "元");
        }
    }

    function replaceProductImage(document, id, imagePath) {
        if (!id || !imagePath) return;
        var imageFile = new File(imagePath);
        if (!imageFile.exists) throw new Error("图片不存在：" + imagePath);
        var original = requireLayer(document, id, "产品图");
        var originalName = original.name;
        var originalGrouped = original.grouped;
        var originalBlendMode = original.blendMode;
        var originalOpacity = original.opacity;
        var originalVisible = original.visible;
        document.activeLayer = original;
        executeAction(stringIDToTypeID("placedLayerMakeCopy"), new ActionDescriptor(), DialogModes.NO);
        var independent = document.activeLayer;
        var descriptor = new ActionDescriptor();
        descriptor.putPath(charIDToTypeID("null"), imageFile);
        executeAction(stringIDToTypeID("placedLayerReplaceContents"), descriptor, DialogModes.NO);
        try { original.remove(); } catch (removeError) { original.visible = false; }
        independent.name = originalName;
        independent.blendMode = originalBlendMode;
        independent.opacity = originalOpacity;
        independent.visible = originalVisible;
        independent.grouped = originalGrouped;
    }

    function applyProduct(document, product) {
        var slot = product.slot;
        var values = product.values;
        setSimpleText(document, slot.fields.selling_point, values.selling_point);
        setSimpleText(document, slot.fields.product_name, values.product_name);
        setSimpleText(document, slot.fields.benefit, values.benefit);
        setSimpleText(document, slot.fields.sales, values.sales);
        setPrice(document, slot, values.price);
        setDiscount(document, slot, values.discount);
        if (values.image) replaceProductImage(document, slot.fields.image, values.image);
    }

    app.displayDialogs = DialogModes.NO;
    log("START total=" + JOBS.length);
    for (var jobIndex = 0; jobIndex < JOBS.length; jobIndex++) {
        var job = JOBS[jobIndex];
        var document = null;
        try {
            log("PAGE_START " + (jobIndex + 1) + "/" + JOBS.length + " " + job.name);
            document = app.open(new File(job.template));
            for (var productIndex = 0; productIndex < job.products.length; productIndex++) {
                log("PRODUCT " + (productIndex + 1) + "/" + job.products.length + " " + job.name);
                applyProduct(document, job.products[productIndex]);
            }
            var output = new File(job.output);
            output.parent.create();
            var options = new PhotoshopSaveOptions();
            options.layers = true;
            options.embedColorProfile = true;
            document.saveAs(output, options, true, Extension.LOWERCASE);
            document.close(SaveOptions.DONOTSAVECHANGES);
            document = null;
            log("PAGE_OK " + job.name + " " + job.output);
        } catch (error) {
            log("PAGE_ERROR " + job.name + " " + error.message + " line=" + error.line);
            try { if (document) document.close(SaveOptions.DONOTSAVECHANGES); } catch (closeError) {}
        }
    }
    log("DONE");
    logFile.close();
})();
'''


def write_batch_script(jobs: list[dict], script_path: str | Path, log_path: str | Path) -> Path:
    script = JSX_ENGINE.replace("__JOBS__", json.dumps(jobs, ensure_ascii=False))
    script = script.replace("__LOG_PATH__", json.dumps(str(Path(log_path).resolve()).replace("\\", "/"), ensure_ascii=False))
    target = Path(script_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(script, encoding="utf-8-sig")
    return target


def launch_batch(photoshop_path: str | Path, script_path: str | Path) -> subprocess.Popen:
    return subprocess.Popen([str(photoshop_path), "-r", str(Path(script_path).resolve())])
