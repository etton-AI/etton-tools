"""
ETTON 提单 + 电放保函 —— 审核流程后端（Flask）
============================================
两步式交互（对应前端"人工审核确认"）：
  ① POST /api/bl/extract   上传报关底单PDF → 提取字段 → 返回 JSON（前端渲染可编辑表格）
  ② POST /api/bl/generate  接收审核确认后的字段数据 → 生成 提单+保函 → 返回 ZIP

核心引擎复用 core.py（提取 / 邮件合并 / docx->PDF / 打包）。
"""
import os, re, json, zipfile, tempfile, subprocess
from datetime import datetime
from flask import Flask, request, jsonify, send_file

import core  # 提取/填充/打包引擎

BASE = os.path.dirname(os.path.abspath(__file__))
UPLOAD = os.path.join(BASE, "workspace", "uploads")
OUT = os.path.join(BASE, "workspace", "output")
MERGED = os.path.join(BASE, "workspace", "merged")
os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUT, exist_ok=True)
os.makedirs(MERGED, exist_ok=True)

BL_TEMPLATE = os.path.join(BASE, "提单模板.docx")

app = Flask(__name__)


@app.errorhandler(Exception)
def handle_unexpected(e):
    """兜底：任何未捕获异常都返回 JSON（而非 Werkzeug 默认的 HTML "Internal Server Error" 页面），
    避免前端 res.json() 报「Unexpected token 'I', "Internal S..." is not valid JSON」。
    同时把 traceback 打到 stderr（pm2 error log 可查真实原因）。"""
    import traceback
    traceback.print_exception(type(e), e, e.__traceback__)
    return jsonify({"ok": False, "error": f"服务器内部错误：{e}"}), 500


def _merged_path(folder):
    """按文件夹名生成稳定的合并底单路径（md5 避免中文/特殊字符路径问题）。"""
    import hashlib
    key = hashlib.md5(str(folder).encode("utf-8")).hexdigest()
    return os.path.join(MERGED, f"{key}.pdf")


# ---------- ⓪ 客户列表（前端下拉渲染，加客户只改 core.CUSTOMERS） ----------
@app.route("/api/bl/customers", methods=["GET"])
def api_customers():
    return jsonify({
        "ok": True,
        "customers": core.list_customers(),
        "default": core.DEFAULT_CUSTOMER,
    })


# ---------- ⓪b 港口/渠道映射读写（前端 /bl-mapping 编辑入口） ----------
def _norm_mapping(d):
    """校验映射表：必须是 {字符串: 字符串}，key 去空白。返回 (是否合法, 归一化 dict)。"""
    if not isinstance(d, dict):
        return False, {}
    out = {}
    for k, v in d.items():
        if not isinstance(k, str) or not isinstance(v, str):
            continue
        k = k.strip()
        v = v.strip()
        if k:
            out[k] = v
    return True, out


@app.route("/api/bl/mappings", methods=["GET"])
def api_get_mappings():
    return jsonify({
        "ok": True,
        "port_map": core.load_port_map(),
        "channel_map": core.load_channel_map(),
        "customs_office_map": core.load_customs_office_map(),
    })


@app.route("/api/bl/mappings", methods=["POST"])
def api_save_mappings():
    payload = request.get_json(force=True)
    pm = payload.get("port_map")
    cm = payload.get("channel_map")
    com = payload.get("customs_office_map")
    if not isinstance(pm, dict) or not isinstance(cm, dict) or not isinstance(com, dict):
        return jsonify({"ok": False, "error": "port_map / channel_map / customs_office_map 必须是对象"}), 400
    ok1, origin = _norm_mapping(pm.get("origin"))
    ok2, destination = _norm_mapping(pm.get("destination"))
    ok3, channel = _norm_mapping(cm)
    ok4, customs_office = _norm_mapping(com)
    if not (ok1 and ok2 and ok3 and ok4):
        return jsonify({"ok": False, "error": "映射的键和值都必须是字符串"}), 400
    core.save_port_map({"origin": origin, "destination": destination})
    core.save_channel_map(channel)
    core.save_customs_office_map(customs_office)
    return jsonify({
        "ok": True,
        "port_map": {"origin": origin, "destination": destination},
        "channel_map": channel,
        "customs_office_map": customs_office,
    })


# ---------- ① 上传报关底单 → 提取字段 → 返回 JSON ----------
@app.route("/api/bl/extract", methods=["POST"])
def api_extract():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "请上传报关底单PDF"}), 400
    f = request.files["file"]
    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "仅支持 PDF 格式的报关底单"}), 400

    # 保存（支持批量：多文件用时间戳区分）
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    saved = os.path.join(UPLOAD, f"{stamp}_{f.filename}")
    f.save(saved)

    customer = (request.form.get("customer") or core.DEFAULT_CUSTOMER).strip()
    try:
        rec = core.extract_customs_data(saved, customer)
    except Exception as e:
        return jsonify({"ok": False, "error": f"提取失败：{e}"}), 500

    # 可选：箱货清单 xlsx → 英文品名 + 总体积（覆盖）+ 收集 FBA ID + 客户渠道
    warning = None
    fba_ids = []
    channel = None
    if "packing" in request.files:
        pf = request.files["packing"]
        if pf and pf.filename:
            p_saved = os.path.join(UPLOAD, f"{stamp}_{pf.filename}")
            pf.save(p_saved)
            try:
                pl = core.parse_packing_list(p_saved)
                if pl.get("products_en"):
                    rec["品名"] = "\n".join(pl["products_en"])
                if pl.get("total_volume"):
                    rec["总体积"] = str(pl["total_volume"])
                if pl.get("so_numbers"):
                    rec["系统SO"] = "\n".join(pl["so_numbers"])
                fba_ids = pl.get("fba_ids", [])
                channel = pl.get("channel")
            except Exception as e:
                warning = f"箱货清单解析失败：{e}"

    # 可选：物流追踪表 xlsx → 按 FBA ID 匹配 ETD/ETA/船名航次
    if "tracking" in request.files:
        tf = request.files["tracking"]
        if tf and tf.filename:
            t_saved = os.path.join(UPLOAD, f"{stamp}_{tf.filename}")
            tf.save(t_saved)
            try:
                track_map = core.parse_tracking_list(t_saved)
                for fid in fba_ids:
                    if fid in track_map:
                        info = track_map[fid]
                        if info.get("etd"):
                            rec["起运日期"] = info["etd"]
                        if info.get("vessel") and not (rec.get("船名航次") or "").strip():
                            rec["船名航次"] = info["vessel"]
                        arrival = core._arrival_date(info.get("etd"), info.get("eta"), express=core._is_express(channel, rec.get("运输方式")))
                        if arrival:
                            rec["申请日期"] = core._date_cn(arrival)
                        break
            except Exception as e:
                msg = f"物流追踪表解析失败：{e}"
                warning = f"{warning}；{msg}" if warning else msg

    # 补全前端表格需要的字段（空值兜底）
    for key in ["总体积"]:
        rec.setdefault(key, "")

    # 港口中文 → 英文（按渠道大类 + 运输方式 + 运抵国）
    core.apply_port_map(rec, channel)
    # 客户渠道 → 目的港英文（非规则渠道兜底，命中则覆盖）
    core.apply_channel_map(rec, channel)

    # 保函「运输工具」为空（非海运单底单该列为 @ 占位符）→ 提示找供应商核实班列号/车次
    if not (rec.get("运输工具") or "").strip():
        msg = "运输工具为空，请手工填写（等于船名航次）"
        warning = f"{warning}；{msg}" if warning else msg

    return jsonify({
        "ok": True,
        "record": rec,
        "bl_fields": core.BL_FIELDS,
        "bl_header": core.BL_HEADER,
        "telex_fields": core.TELEX_FIELDS,
        "warning": warning,
    })


# ---------- ①b 批量提取：按文件夹区分，多票一次上传 ----------
@app.route("/api/bl/extract-batch", methods=["POST"])
def api_extract_batch():
    """批量上传：同一票的底单 PDF（可多份拆分）+ 箱货清单在同一文件夹；物流追踪表全局一份。

    FormData 字段约定（前端逐个 append）：
      folder_{i}       第 i 票的文件夹名（用于命名 + 箱数/国家校验）
      pdf_{i}_{j}      第 i 票的第 j 份报关底单 PDF（1 份或多份拆分）
      packing_{i}      第 i 票的箱货清单 xlsx（可选）
      tracking         物流追踪表 xlsx（可选，全局一份）
    """
    # 1. 解析 FormData
    tickets = {}  # idx -> {folder, pdfs: [FileStorage], packing: FileStorage}
    for key, f in request.files.items(multi=True):
        m = re.match(r"pdf_(\d+)_(\d+)", key)
        if m:
            idx = int(m.group(1))
            t = tickets.setdefault(idx, {"folder": "", "pdfs": [], "packing": None})
            t["pdfs"].append(f)
            continue
        m = re.match(r"packing_(\d+)", key)
        if m:
            idx = int(m.group(1))
            t = tickets.setdefault(idx, {"folder": "", "pdfs": [], "packing": None})
            t["packing"] = f
    for key, v in request.form.items(multi=True):
        m = re.match(r"folder_(\d+)", key)
        if m:
            idx = int(m.group(1))
            t = tickets.setdefault(idx, {"folder": "", "pdfs": [], "packing": None})
            t["folder"] = v

    if not tickets:
        return jsonify({"ok": False, "error": "未收到任何票的报关底单"}), 400

    # 2. 物流追踪表：解析一次，全局共享（{FBA ID: {etd, eta, vessel}}）
    track_map = {}
    if "tracking" in request.files:
        tf = request.files["tracking"]
        if tf and tf.filename:
            tf_path = os.path.join(UPLOAD, f"tracking_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.xlsx")
            tf.save(tf_path)
            try:
                track_map = core.parse_tracking_list(tf_path)
            except Exception:
                track_map = {}

    # 2b. 周汇总箱货清单：解析一次，全局共享（按工作号分组，供逐票匹配）
    weekly_groups = []
    if "weekly_packing" in request.files:
        wf = request.files["weekly_packing"]
        if wf and wf.filename:
            w_path = os.path.join(UPLOAD, f"weekly_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.xlsx")
            wf.save(w_path)
            try:
                weekly_groups = core.parse_packing_list_weekly(w_path)
            except Exception:
                weekly_groups = []

    # 3. 逐票处理
    customer = (request.form.get("customer") or core.DEFAULT_CUSTOMER).strip()
    out_tickets = []
    for idx in sorted(tickets):
        t = tickets[idx]
        folder = t["folder"]
        pdfs = t["pdfs"]
        if not pdfs:
            continue
        stamp = datetime.now().strftime("%Y%m%d%H%M%S%f")

        # 保存所有 PDF，合并成一份（单份也统一走合并）
        # 注意：保留原始文件名，merge 排序靠文件名里的「报/放/委托」关键词识别拆分底单顺序
        saved_paths = []
        for j, pf in enumerate(pdfs):
            safe_name = re.sub(r'[\\/:*?"<>|]', "_", pf.filename or f"pdf{j}.pdf")
            p = os.path.join(UPLOAD, f"batch_{idx}_{j}_{stamp}_{safe_name}")
            pf.save(p)
            saved_paths.append(p)
        merged_pdf = _merged_path(folder)
        try:
            core.merge_declaration_pdfs(saved_paths, merged_pdf)
        except Exception:
            # pypdf 合并失败（如加密/损坏）时退回第一份
            import shutil as _sh
            _sh.copy(saved_paths[0], merged_pdf)

        # 提取字段
        try:
            rec = core.extract_customs_data(merged_pdf, customer)
        except Exception as e:
            return jsonify({"ok": False, "error": f"票「{folder}」提取失败：{e}"}), 500
        # 「文件命名」= 上传文件夹名（ZIP 内底单/提单/保函的文件名均由该文件夹名派生，而非提单号）
        rec["文件命名"] = folder

        # 箱货清单：优先周汇总匹配，匹配不到回退每票单票清单
        warning = None
        fba_ids = []
        channel = None
        packing_meta = None
        if weekly_groups:
            wm = core.match_weekly_packing(rec, weekly_groups)
            if wm:
                packing_meta = wm
                if wm.get("products_en"):
                    rec["品名"] = "\n".join(wm["products_en"])
                if wm.get("total_volume"):
                    rec["总体积"] = str(wm["total_volume"])
                if wm.get("so_numbers"):
                    rec["系统SO"] = "\n".join(wm["so_numbers"])
                fba_ids = wm.get("fba_ids", [])
                channel = wm.get("channel")
        if packing_meta is None and t["packing"] and t["packing"].filename:
            p_saved = os.path.join(UPLOAD, f"packing_{idx}_{stamp}.xlsx")
            t["packing"].save(p_saved)
            try:
                pl = core.parse_packing_list(p_saved)
                packing_meta = pl
                if pl.get("products_en"):
                    rec["品名"] = "\n".join(pl["products_en"])
                if pl.get("total_volume"):
                    rec["总体积"] = str(pl["total_volume"])
                if pl.get("so_numbers"):
                    rec["系统SO"] = "\n".join(pl["so_numbers"])
                fba_ids = pl.get("fba_ids", [])
                channel = pl.get("channel")
            except Exception as e:
                warning = f"箱货清单解析失败：{e}"

        # 物流追踪表：按 FBA ID 匹配 ETD/ETA/船名航次
        for fid in fba_ids:
            if fid in track_map:
                info = track_map[fid]
                if info.get("etd"):
                    rec["起运日期"] = info["etd"]
                if info.get("vessel") and not (rec.get("船名航次") or "").strip():
                    rec["船名航次"] = info["vessel"]
                arrival = core._arrival_date(info.get("etd"), info.get("eta"), express=core._is_express(channel, rec.get("运输方式")))
                if arrival:
                    rec["申请日期"] = core._date_cn(arrival)
                break

        # 港口/渠道映射（渠道大类 + 运输方式 + 运抵国）
        core.apply_port_map(rec, channel)
        core.apply_channel_map(rec, channel)

        # 保函「运输工具」为空提醒
        if not (rec.get("运输工具") or "").strip():
            msg = "运输工具为空，请手工填写（等于船名航次）"
            warning = f"{warning}；{msg}" if warning else msg

        # 国家/箱数一致性校验（底单 vs 箱货清单 vs 文件夹名）
        for w in core.validate_ticket(folder, rec, packing_meta):
            warning = f"{warning}；{w}" if warning else w

        out_tickets.append({"folder": folder, "record": rec, "warning": warning})

    if not out_tickets:
        return jsonify({"ok": False, "error": "无有效票"}), 400

    return jsonify({
        "ok": True,
        "tickets": out_tickets,
        "bl_fields": core.BL_FIELDS,
        "bl_header": core.BL_HEADER,
        "telex_fields": core.TELEX_FIELDS,
    })


# ---------- ② 接收审核后数据 → 生成 提单+保函 → 返回 ZIP ----------
@app.route("/api/bl/generate", methods=["POST"])
def api_generate():
    payload = request.get_json(force=True)
    tickets = payload.get("tickets")
    customer = (payload.get("customer") or core.DEFAULT_CUSTOMER).strip()
    root_folder = (payload.get("root_folder") or "").strip()
    if not tickets:
        # 兼容旧版单票 records 调用
        records = payload.get("records") or [payload.get("record")]
        tickets = [{
            "folder": (r.get("文件命名") or r.get("提单号") or "提单") if isinstance(r, dict) else "提单",
            "record": r,
            "merged_pdf": None,
        } for r in records if isinstance(r, dict)]
    if not tickets:
        return jsonify({"ok": False, "error": "无数据"}), 400

    # 记忆回写：中文港口 → 英文；带渠道的票改目的港后记入 channel_map
    records = [t["record"] for t in tickets]
    core.remember_ports(records)
    core.remember_channels(records)

    # 补充每票的合并底单路径（按文件夹名重建）
    for t in tickets:
        if not t.get("merged_pdf"):
            t["merged_pdf"] = _merged_path(t.get("folder") or "")

    out_dir = os.path.join(OUT, datetime.now().strftime("%Y%m%d%H%M%S"))
    try:
        zip_path, results = core.generate_batch(tickets, BL_TEMPLATE, out_dir, customer, root_folder=root_folder)
    except Exception as e:
        return jsonify({"ok": False, "error": f"生成失败：{e}"}), 500

    # 返回 ZIP 下载相对路径 + 每票预览相对路径（前端经 /api/bl/file/<rel> 访问）
    def _rel(p):
        return os.path.relpath(p, OUT).replace("\\", "/")

    previews = [{
        "folder": r["folder"],
        "bl": _rel(r["bl_pdf"]),
        "telex": _rel(r["telex_preview_pdf"]),
    } for r in results]

    return jsonify({"ok": True, "zip": _rel(zip_path), "previews": previews})


# ---------- ③ 生成文件服务：预览 / 下载 ----------
@app.route("/api/bl/file/<path:rel>")
def api_file(rel):
    base = os.path.abspath(OUT)
    full = os.path.abspath(os.path.join(base, rel))
    if full != base and not full.startswith(base + os.sep):
        return jsonify({"ok": False, "error": "非法路径"}), 403
    if not os.path.isfile(full):
        return jsonify({"ok": False, "error": "文件不存在"}), 404
    return send_file(full)


if __name__ == "__main__":
    # pm2 生产环境关 debug（避免 Werkzeug reloader 在 pm2 下反复重启）；
    # 本地手动调试可 `FLASK_DEBUG=1 python review_app.py`
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=5000, debug=debug)
