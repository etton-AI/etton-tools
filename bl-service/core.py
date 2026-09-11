"""
ETTON 提单 + 电放保函 生成引擎（审核流程版）
============================================
流程：报关底单PDF -> 提取字段 -> 生成「审核用Excel」(供前端展示/人工编辑)
      -> 用户确认后 -> 邮件合并生成 提单.docx + 电放保函.docx -> ZIP打包

关键设计：
1. 提单模板使用真实 ETTON 模板的 Word 邮件合并域（MERGEFIELD），直接填域，保格式不漂移。
2. 电放保函按真实保函结构生成。
3. 审核Excel = 提单信息收集表，字段列固定，每行一票。
"""
import os, re, json, zipfile, shutil, subprocess, tempfile
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------- 客户固定信息（shipper/consignee 英文，来自正确提单/保函） ----------
# 每个客户一套固定信息；提单/保函按客户分发。新增客户在此加一条配置即可。
CUSTOMERS = {
    "拓锐": {
        "label": "拓锐（广州拓锐科技有限公司）",
        "shipper_bl": (
            "GUANGZHOU TUORUI TECHNOLOGY CO., LTD\n"
            "ROOM 411, NO.101. DEXING ROAD WANGGANG JIAHE\n"
            "STREET BAIYUN DISTRICT GUANGZHOU"
        ),
        "consignee_bl": "HONG KONG LIXIANG TRADING COMPANY LIMITED",
        "shipper_telex": "Guangzhou Tuorui Technology Co.,Ltd",
        "consignee_telex": "Hong Kong Lixiang Trading Company Limited",
        # 保函页眉 / TO / FROM（来自拓锐标准答案电放保函；TO 保留「Co. ,Ltd」写法）
        "telex_header": "广州拓锐科技有限公司",
        "telex_to": "Guangzhou Tuorui Technology Co. ,Ltd",
        "telex_from": "ETTON TECHNOLOGY LOGISTICS (ZHONGSHAN) CO., LTD",
    },
    "星速": {
        "label": "星速（HNXS）",
        # 星速为 Amazon FBA 直发：收货人固定 Amazon，通知方 = 同收货人；
        # 发货人可变（境内发货人 → 拼音大写）；无需电放保函（no_telex）。
        "shipper_bl": None,  # None = 可变发货人（境内发货人拼音），extract 里单独处理
        "consignee_bl": "AMAZONFULFILMENTCENTER",
        "notify_bl": "SAMEASCONSIGNEE",
        "shipper_telex": "",
        "consignee_telex": "",
        "telex_header": "",
        "telex_to": "",
        "telex_from": "",
        "no_telex": True,
    },
}
DEFAULT_CUSTOMER = "拓锐"


def get_customer_config(customer):
    """取客户配置，未知客户回退默认（拓锐）。"""
    return CUSTOMERS.get(customer, CUSTOMERS[DEFAULT_CUSTOMER])


def list_customers():
    """返回客户列表 [{key, label}]，供前端下拉渲染（加客户只改 CUSTOMERS 一处）。"""
    return [{"key": k, "label": v["label"]} for k, v in CUSTOMERS.items()]


def _cn_to_pinyin_upper(name):
    """中文公司名 → 拼音大写无空格（星速提单 shipper 格式）。

    例：湖南永高商贸有限公司 → HUNANYONGGAOSHANGMAOYOUXIANGONGSI。
    含括号（如公司名带 (9143...) 统一社会信用代码）时先去括号内容；pypinyin 不可用则原样返回。
    """
    name = re.sub(r"[\(（][^)）]*[\)）]", "", (name or "").strip())
    if not name:
        return ""
    try:
        from pypinyin import lazy_pinyin
        return "".join(lazy_pinyin(name)).upper()
    except Exception:
        return name.upper()


def _base_fba(fid):
    """箱货清单 FBA ID → 基础 FBA（12 位 FBA+9 位，去「U+6位序列号」后缀）。

    例：FBA15LHBVDHZU000001 → FBA15LHBVDHZ；FBA15L9QBCCTU000001 → FBA15L9QBCCT。
    底单/订单列表里的 FBA 就是 12 位基础 FBA，两边一致即可匹配。
    """
    fid = (fid or "").strip()
    m = re.match(r"^(FBA\d{2}[0-9A-Z]{7})", fid)
    return m.group(1) if m else fid


def _fba_from_text(text):
    """从任意文本里提取基础 FBA（如 合同协议号 FBA15LHCTRGN+HCXX5P → FBA15LHCTRGN）。"""
    m = re.search(r"FBA\d{2}[0-9A-Z]{7}", text or "")
    return m.group(0) if m else ""


def _fbas_from_text(text):
    """提取文本里的全部基础 FBA（含 '+6位后缀' 缩写，如 FBA15M0WQDJZ+0S50DJ → [FBA15M0WQDJZ, FBA15M0S50DJ]）。

    底单文件名里的多 FBA 有两种写法：逗号分隔完整 FBA（FBA15LZW7KR4,FBA15M03GFHD），
    或 '+6位后缀' 缩写（FBA15M0WQDJZ+0S50DJ，第二个 FBA 与前一个共享前 6 位 FBA15M）。
    """
    found = []
    text = text or ""
    for m in re.finditer(r"FBA\d{2}[0-9A-Z]{7}", text):
        base = m.group(0)
        if base not in found:
            found.append(base)
        # '+6位后缀'：第二个 FBA 与前一个共享前 6 位（如 FBA15M），还原为完整 FBA
        m2 = re.match(r"\+([0-9A-Z]{6})", text[m.end():])
        if m2:
            full = base[:6] + m2.group(1)
            if full not in found:
                found.append(full)
    return found

# 中文品名 -> 英文品名（来自箱货清单，可扩展）
PRODUCT_EN_MAP = {
    "折叠台灯": "Foldable Desk lamp",
    "魔法球投影灯": "Magic Ball Night Light Projector",
    "圆饼夹子灯": "Cilp Table Lamp",
    "水波纹灯": "INS Ripple Lights Projector",
    "玫瑰投影灯": "Rose Projector Light",
    "旋转小屋投影灯": "Princess Night Light Projector",
    "冰川极光投影灯": "Glacial Northern Lights Projector",
    "恐龙蛋投影灯": "Dinosuar Egg Night Light Projector2.0",
}

# ---------- 港口中英对照（port_map.json，可随日常使用自动积累） ----------
PORT_MAP_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "port_map.json")


def load_port_map():
    """读取港口中英对照表，结构 {origin: {中文口岸: 英文}, destination: {中文运抵国: 英文}}。
    文件不存在或损坏返回空结构。兼容旧版扁平格式：无 origin/destination 键时全量归 destination（历史主用途为运抵国映射）。"""
    empty = {"origin": {}, "destination": {}}
    try:
        with open(PORT_MAP_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return empty
        if "origin" in data or "destination" in data:
            return {
                "origin": data.get("origin") or {},
                "destination": data.get("destination") or {},
            }
        # 旧版扁平格式：{中文: 英文}
        return {"origin": {}, "destination": data}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return empty


def save_port_map(pm):
    """保存港口中英对照表到 port_map.json。"""
    try:
        with open(PORT_MAP_PATH, "w", encoding="utf-8") as f:
            json.dump(pm, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def _is_english(s):
    """字符串是否不含中文（用于港口记忆：只把英文映射记入对照表）。"""
    return not re.search(r"[一-鿿]", s or "")


# ---------- 渠道大类规则（客户渠道关键词 → 铁路/卡航/快递） ----------
# 渠道大类来自箱货清单「客户渠道」列，与「运输方式」（底单提取）是两套维度：
# 目的港/起运地按渠道大类分流（铁路/卡航/快递各有固定规则）。
_CHANNEL_CATEGORY_KEYWORDS = [
    ("铁路", ["铁路", "快铁", "铁派"]),
    ("卡航", ["卡航"]),
    ("快递", ["联邦", "空派"]),
]
# 渠道大类 → 目的港规则（按「运抵国」查）
_ROUTE_DEST_RULES = {
    "铁路": {"英国": "MALASZEWICZE", "德国": "DUISBURG"},
    "卡航": {"英国": "THE UK", "德国": "GERMANY"},
}
# 快递渠道（香港联邦IP）固定起运地
_EXPRESS_ORIGIN = "SHENZHEN"


def _channel_category(channel):
    """客户渠道名 → 大类（铁路/卡航/快递/None）。按关键词顺序匹配，先命中先返回。"""
    ch = (channel or "").strip()
    if not ch:
        return None
    for cat, kws in _CHANNEL_CATEGORY_KEYWORDS:
        for kw in kws:
            if kw in ch:
                return cat
    return None


def _is_express(channel, mode):
    """是否「快」渠道（空运/国际快递）——这类运输很快，缺 ETA 时不做 ETD+10 兜底。
    海运/铁路/卡航等慢渠道才适用 ETD+10。"""
    if (mode or "").strip() == "航空运输":
        return True
    if _channel_category(channel) == "快递":
        return True
    return False


def apply_port_map(rec, channel=None):
    """把 rec 里的「起运地/目的港」按对照表/渠道规则转英文；同时记录原始值供记忆回写。
    起运港：
      - 快递（香港联邦IP）→ 固定 SHENZHEN
      - 海运（水路运输）→ 离境口岸 → port_map.origin
      - 非海运（铁路/公路/航空）→ 关区名 → customs_office_map
    目的港（按「运抵国」查，非「指运港」）：
      - 铁路 → {英国:MALASZEWICZE, 德国:DUISBURG}
      - 卡航 → {英国:THE UK, 德国:GERMANY}
      - 其他 → port_map.destination。"""
    pm = load_port_map()
    origins = pm.get("origin", {}) or {}
    dests = pm.get("destination", {}) or {}
    offices = load_customs_office_map()

    cat = _channel_category(channel)
    mode = (rec.get("运输方式") or "").strip()
    office = (rec.get("关区名") or "").strip()
    cn = (rec.get("起运地") or "").strip()   # extract 里「起运地」= 离境口岸（中文）
    if cat == "快递":
        rec["_起运地_原始"] = channel
        rec["起运地"] = _EXPRESS_ORIGIN
        rec["_起运地_固定"] = True
    elif mode == "水路运输":
        rec["_起运地_原始"] = cn
        if cn and cn in origins:
            rec["起运地"] = origins[cn]
    else:
        # 非海运：起运港取关区名（关区名缺失时保留离境口岸中文供人工填）
        rec["_起运地_原始"] = office
        if office and office in offices:
            rec["起运地"] = offices[office]

    country = (rec.get("运抵国") or "").strip()
    rec["_目的港_原始"] = country
    dest = None
    if cat in _ROUTE_DEST_RULES and country in _ROUTE_DEST_RULES[cat]:
        dest = _ROUTE_DEST_RULES[cat][country]
    elif country in dests:
        dest = dests[country]
    if dest:
        rec["目的港"] = dest
    # 保函「目的地」= 目的港，同步英文
    if rec.get("目的港"):
        rec["目的地"] = rec.get("目的港", "")
    return rec


def remember_ports(records):
    """记忆回写：对比「原始中文」与「用户最终英文」，把新映射追加进对照表。
    起运地按运输方式分流：海运 → port_map.origin（离境口岸），非海运 → customs_office_map（关区名）；
    目的港 → port_map.destination。返回是否新增。"""
    pm = load_port_map()
    origins = pm.setdefault("origin", {})
    dests = pm.setdefault("destination", {})
    offices = load_customs_office_map()
    changed = False
    office_changed = False
    for rec in records:
        mode = (rec.get("运输方式") or "").strip()
        cn = (rec.get("_起运地_原始") or "").strip()
        en = (rec.get("起运地") or "").strip()
        # 快递渠道起运地固定 SHENZHEN，无需记忆回写（_起运地_固定 标记）
        if not rec.get("_起运地_固定") and cn and en and en != cn and _is_english(en):
            if mode == "水路运输":
                if cn not in origins:
                    origins[cn] = en
                    changed = True
            else:
                if cn not in offices:
                    offices[cn] = en
                    office_changed = True
        cn = (rec.get("_目的港_原始") or "").strip()
        en = (rec.get("目的港") or "").strip()
        if cn and en and en != cn and cn not in dests and _is_english(en):
            dests[cn] = en
            changed = True
    if changed:
        save_port_map(pm)
    if office_changed:
        save_customs_office_map(offices)
    return changed or office_changed


# ---------- 关区名 → 起运港对照（customs_office_map.json，非海运票用） ----------
# 非海运（铁路/公路卡航/航空）提单的起运港不取「离境口岸」（那是边境口岸），
# 而是取报关单「海关编号」后面的关区名（报关关区 ≈ 发货城市），再查此表转英文。
CUSTOMS_OFFICE_MAP_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "customs_office_map.json")


def load_customs_office_map():
    """读取关区名→英文起运港对照表 {关区名: 英文}。文件不存在或损坏返回空 dict。"""
    try:
        with open(CUSTOMS_OFFICE_MAP_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_customs_office_map(om):
    """保存关区名→英文起运港对照表到 customs_office_map.json。"""
    try:
        with open(CUSTOMS_OFFICE_MAP_PATH, "w", encoding="utf-8") as f:
            json.dump(om, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def _customs_office_from_text(text):
    """从报关单文本提取「海关编号」后面的关区名（如 增城海关 / 蓉青关 / 深圳湾关）。
    格式：`海关编号：123... (增城海关)` 或 `海关编号：123... (7901) 蓉青关`。"""
    m = re.search(r'海关编号[：:]?\s*(\d+)\s*[\(（]([^\)）]*)[\)）]', text)
    if not m:
        return ""
    inner = m.group(2).strip()
    if re.search(r'[一-鿿]', inner):      # 括号内是关区名
        return inner
    rest = text[m.end():]                 # 括号内是关区代码 → 关区名在括号后
    n = re.search(r'([一-鿿][一-鿿关海]*关?)', rest)
    return n.group(1) if n else ""


# ---------- 渠道 → 目的港对照（channel_map.json，与箱货清单「客户渠道」关联） ----------
CHANNEL_MAP_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "channel_map.json")


def load_channel_map():
    """读取渠道→目的港对照表 {渠道名: 英文目的港}。文件不存在或损坏返回空 dict。"""
    try:
        with open(CHANNEL_MAP_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_channel_map(cm):
    """保存渠道→目的港对照表到 channel_map.json。"""
    try:
        with open(CHANNEL_MAP_PATH, "w", encoding="utf-8") as f:
            json.dump(cm, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def apply_channel_map(rec, channel):
    """按箱货清单「客户渠道」查 channel_map 覆盖目的港英文（比中文港口名更精确）。
    仅对非铁路/卡航/快递渠道生效——这三类已由 _ROUTE_DEST_RULES 规则化，避免被旧 channel_map 覆盖。
    未命中保留原值；同时记录原始渠道供记忆回写。"""
    ch = (channel or "").strip()
    rec["_目的港_渠道"] = ch
    if _channel_category(ch) is not None:
        return rec
    cm = load_channel_map()
    if ch and ch in cm:
        rec["目的港"] = cm[ch]
        if rec.get("目的港"):
            rec["目的地"] = rec["目的港"]
    return rec


def remember_channels(records):
    """记忆回写：用户把目的港改成英文后，若该票带渠道，则把「渠道→英文目的港」记入 channel_map.json。
    铁路/卡航/快递渠道规则化，不按渠道名记忆。返回是否新增。"""
    cm = load_channel_map()
    changed = False
    for rec in records:
        ch = (rec.get("_目的港_渠道") or "").strip()
        en = (rec.get("目的港") or "").strip()
        if _channel_category(ch) is not None:
            continue
        if ch and en and ch not in cm and _is_english(en):
            cm[ch] = en
            changed = True
    if changed:
        save_channel_map(cm)
    return changed


# ---------- 星速（HNXS）专用：目的港 / 起运港 / 订单列表 / 箱货清单 ----------
# 星速为 Amazon FBA 直发，无需电放保函（shipper=境内发货人拼音，
# consignee 固定 Amazon，通知方=同收货人）。提单与拓锐共用同一 ETTON 邮件合并模板。
# 目的港按「发往国家 + 运输方式（海运/陆运）」映射。
XS_DEST_MAP_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xs_dest_map.json")

# 默认目的港映射（可随使用自动积累）：陆运取国家英文名，海运取实际港口
XS_DEST_MAP_DEFAULT = {
    "海运": {"英国": "FELIXSTOWE", "德国": "ROTTERDAM,NL", "法国": "ROTTERDAM,NL", "荷兰": "ROTTERDAM,NL"},
    "陆运": {"英国": "BRITAIN", "德国": "GERMANY", "法国": "FRANCE", "荷兰": "NETHERLANDS"},
}


def load_xs_dest_map():
    """读取星速目的港映射 {海运: {国家: 港口}, 陆运: {国家: 港口}}。缺失时回退默认值。"""
    data = None
    try:
        with open(XS_DEST_MAP_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        data = None
    if not isinstance(data, dict):
        data = dict(XS_DEST_MAP_DEFAULT)
    for key in ("海运", "陆运"):
        data.setdefault(key, dict(XS_DEST_MAP_DEFAULT.get(key, {})))
    return data


def save_xs_dest_map(dm):
    """保存星速目的港映射到 xs_dest_map.json。"""
    try:
        with open(XS_DEST_MAP_PATH, "w", encoding="utf-8") as f:
            json.dump(dm, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def apply_xs_map(rec, order_info):
    """星速提单字段映射：起运港（海运→离境口岸查 port_map.origin，其余→SHENZHEN）、
    目的港（发往国家 + 运输方式查 xs_dest_map；加拿大美转加按供应渠道分 LONGBEACH/LOSANGELES）、
    起运日期（订单列表开船时间）、海运船名航次（船名/船次）。
    order_info 为订单列表匹配到的 {country, ship_time, vessel, voyage, biz_type, supply, channel}。
    """
    order_info = order_info or {}
    mode = (rec.get("运输方式") or "").strip()
    biz = (order_info.get("biz_type") or "").strip()
    is_sea = "水" in mode or biz == "海运"

    # 起运地：海运取「离境口岸」（extract 已把离境口岸存进 rec['起运地']）→ port_map.origin；
    #        非海运（卡航/铁路/公路）固定 SHENZHEN。
    pm = load_port_map()
    origins = pm.get("origin", {}) or {}
    cn_port = (rec.get("起运地") or "").strip()
    rec["_起运地_原始"] = cn_port
    if is_sea:
        if cn_port and cn_port in origins:
            rec["起运地"] = origins[cn_port]
        # 海运但离境口岸无对照 → 保留中文供人工填
    else:
        rec["起运地"] = "SHENZHEN"

    # 目的港
    country = (order_info.get("country") or "").strip() or (rec.get("运抵国") or "").strip()
    rec["_目的港_原始"] = country
    supply = (order_info.get("supply") or "").strip() or (order_info.get("channel") or "").strip()
    dest = None
    if "加拿大" in country:
        # 美转加（经美中转）：带美森/CLX → LONGBEACH,CA；非美森 → LOSANGELES,CA。
        # 直航加拿大（加东/加西普船）目的港是加拿大本土港口，留空人工填。
        if "美转加" in supply:
            dest = "LONGBEACH,CA" if ("美森" in supply or "CLX" in supply.upper()) else "LOSANGELES,CA"
    elif country:
        dm = load_xs_dest_map()
        dest = dm.get("海运" if is_sea else "陆运", {}).get(country)
    if dest:
        rec["目的港"] = dest
        rec["目的地"] = dest

    # 星速提单「起运日期」= 订单列表「开船时间」（非报关单申报日期），空则留空人工填
    ship_time = (order_info.get("ship_time") or "").strip()
    rec["起运日期"] = ship_time

    # 海运船名航次：订单列表船名/船次优先，缺省保留报关单提取的运输工具
    if is_sea and not (rec.get("船名航次") or "").strip():
        vessel = (order_info.get("vessel") or "").strip()
        voyage = (order_info.get("voyage") or "").strip()
        if vessel:
            rec["船名航次"] = f"{vessel}/{voyage}" if voyage else vessel
    return rec


def xs_apply_packing(rec, folder, xs_packing, order_map):
    """星速：按文件夹名里的全部 FBA 匹配箱货清单 + 订单列表，回填提单字段。

    - 品名：箱货清单英文品名，跨全部 FBA 去重合并（提单每个品名一行）
    - 总体积：订单列表「总CBM」，按唯一订单行（系统SO 去重）求和（'FBA15M0WQDJZ+0S50DJ' 为两单合并）
    - 系统SO：跨全部 FBA 去重
    - 其余（发往国家/开船时间/供应渠道/船名航次/起运地/目的港）由 apply_xs_map 用首个匹配订单行映射
    """
    fbas = _fbas_from_text(folder)
    if not fbas:
        f = (rec.get("FBA") or "").strip()
        if f:
            fbas = [f]

    products = []
    sos = []
    seen_orders = {}  # 系统SO -> order_info（体积按唯一订单行求和，避免逗号多 FBA 同一行重复累加）
    order_info = {}
    for fba in fbas:
        pk = xs_packing.get(fba)
        if pk:
            for p in pk.get("products_en", []):
                if p and p not in products:
                    products.append(p)
            if pk.get("so") and pk["so"] not in sos:
                sos.append(pk["so"])
        oi = order_map.get(fba)
        if oi:
            if not order_info:
                order_info = oi
            so = (oi.get("so") or "").strip() or fba
            if so not in seen_orders:
                seen_orders[so] = oi

    if products:
        rec["品名"] = "\n".join(products)
    if sos:
        rec["系统SO"] = "\n".join(sos)

    # 总体积：订单列表「总CBM」按唯一订单行求和
    cbm_sum = 0.0
    has_cbm = False
    for oi in seen_orders.values():
        c = oi.get("cbm")
        if isinstance(c, (int, float)):
            cbm_sum += float(c)
            has_cbm = True
    if has_cbm:
        rec["总体积"] = str(round(cbm_sum, 4))

    apply_xs_map(rec, order_info)
    return rec


def parse_order_list(xlsx_path):
    """解析星速「订单列表」xlsx，返回 {基础 FBA: {country, ship_time, vessel, voyage, biz_type, so, supply, channel, cbm, kg, boxes}}。

    表头按文字定位列：发往国家 / 开船时间 / 船名 / 船次 / 业务类型 / 系统SO / 供应渠道 / 客户渠道 / 总CBM / 总KG / 总箱数 / FBA。
    FBA 列可能一个订单含多个 FBA（逗号/换行分隔），逐个拆分映射到同一订单信息。
    """
    # 注意：订单列表用 read_only=False（该文件 read_only 模式会漏读数据行，仅返回表头）
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # 表头文字 -> info 键（用「文字→键」映射避免逐个 if 分支）
    _WANT = {
        "FBA": "fba", "发往国家": "country", "开船时间": "ship_time", "船名": "vessel",
        "船次": "voyage", "业务类型": "biz_type", "系统SO": "so", "供应渠道": "supply",
        "客户渠道": "channel", "总CBM": "cbm", "总KG": "kg", "总箱数": "boxes",
    }
    cols = {}
    header_idx = None
    for i, row in enumerate(rows[:5]):
        if not row:
            continue
        for j, v in enumerate(row):
            s = str(v).strip() if v is not None else ""
            if s in _WANT and _WANT[s] not in cols:
                cols[_WANT[s]] = j
        if "fba" in cols and header_idx is None:
            header_idx = i
    if "fba" not in cols:
        return {}

    fba_col = cols["fba"]

    def _g(row, key):
        c = cols.get(key)
        return str(row[c]).strip() if c is not None and c < len(row) and row[c] else ""

    def _num(row, key):
        c = cols.get(key)
        if c is None or c >= len(row) or row[c] is None:
            return ""
        return float(row[c]) if isinstance(row[c], (int, float)) else ""

    index = {}
    for row in rows[(header_idx + 1):]:
        if not row or fba_col >= len(row) or not row[fba_col]:
            continue
        info = {
            "country": _g(row, "country"),
            "ship_time": _excel_date_str(row[cols["ship_time"]]) if "ship_time" in cols and cols["ship_time"] < len(row) and row[cols["ship_time"]] else "",
            "vessel": _g(row, "vessel"),
            "voyage": _g(row, "voyage"),
            "biz_type": _g(row, "biz_type"),
            "so": _g(row, "so"),
            "supply": _g(row, "supply"),
            "channel": _g(row, "channel"),
            "cbm": _num(row, "cbm"),
            "kg": _num(row, "kg"),
            "boxes": _num(row, "boxes"),
        }
        # 一个订单可能多 FBA（逗号/换行/分号分隔）
        for fid in re.split(r"[,\n，、;；]+", str(row[fba_col])):
            fid = _base_fba(fid)
            if fid:
                index.setdefault(fid, info)
    return index


def parse_packing_list_xs(xlsx_path):
    """解析星速「箱货清单」xlsx，按基础 FBA 分组，返回 {基础FBA: 汇总}。

    每项：products_en(去重英文品名), total_volume(长×宽×高×箱数求和, CBM),
          total_weight(货箱重量求和), boxes(总箱数求和), country, channel, so。
    星速箱货清单无「体积」列，由 长(CM)×宽(CM)×高(CM)×总箱数 折算。
    """
    wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    fba_col = en_col = wt_col = box_col = None
    len_col = wid_col = hei_col = country_col = channel_col = so_col = None
    header_idx = None
    for i, row in enumerate(rows[:12]):
        if not row:
            continue
        for j, v in enumerate(row):
            s = str(v).strip() if v is not None else ""
            if s == "FBA ID":
                fba_col = j
                header_idx = i
            elif s == "英文品名":
                en_col = j
            elif s == "货箱重量":
                wt_col = j
            elif s == "总箱数(CTN)":
                box_col = j
            elif s == "长(CM)":
                len_col = j
            elif s == "宽(CM)":
                wid_col = j
            elif s == "高(CM)":
                hei_col = j
            elif s in ("国家", "目的国"):
                country_col = j
            elif s == "客户渠道":
                channel_col = j
            elif s == "系统SO":
                so_col = j

    groups = {}
    for i, row in enumerate(rows):
        if not row:
            continue
        if header_idx is not None and i <= header_idx:
            continue
        if fba_col is None or fba_col >= len(row) or not row[fba_col]:
            continue
        base = _base_fba(str(row[fba_col]))
        if not base:
            continue
        g = groups.setdefault(base, {
            "products_en": [], "total_volume": 0.0, "total_weight": 0.0,
            "boxes": 0, "country": None, "channel": None, "so": None,
        })
        if so_col is not None and so_col < len(row) and row[so_col] and g["so"] is None:
            g["so"] = str(row[so_col]).strip()
        if country_col is not None and country_col < len(row) and row[country_col] and g["country"] is None:
            g["country"] = str(row[country_col]).strip()
        if channel_col is not None and channel_col < len(row) and row[channel_col] and g["channel"] is None:
            g["channel"] = str(row[channel_col]).strip()
        if en_col is not None and en_col < len(row) and row[en_col]:
            e = _strip_version(str(row[en_col])).upper().replace(" ", "")
            if e and e not in g["products_en"]:
                g["products_en"].append(e)
        if wt_col is not None and wt_col < len(row) and isinstance(row[wt_col], (int, float)):
            g["total_weight"] += float(row[wt_col])
        if box_col is not None and box_col < len(row) and isinstance(row[box_col], (int, float)):
            g["boxes"] += int(row[box_col])
        # 体积：长×宽×高×箱数 / 1e6（星速箱货清单无「体积」列）
        if None not in (len_col, wid_col, hei_col) and max(len_col, wid_col, hei_col) < len(row):
            L, W, H = row[len_col], row[wid_col], row[hei_col]
            B = row[box_col] if box_col is not None and box_col < len(row) and isinstance(row[box_col], (int, float)) else 1
            if all(isinstance(x, (int, float)) for x in (L, W, H)):
                g["total_volume"] += float(L) * float(W) * float(H) * float(B) / 1e6
    return groups


# ---------- 1. 报关单字段提取（标签定位版，适配横向 842x595 报关单） ----------
# 报关底单页面的关键标签（用于判断某页/某 PDF 是否是报关底单，跳过「委托报关协议」等非底单）
_BL_KEY_LABELS = ("境内发货人", "境外收货人", "提运单号", "件数", "毛重", "离境口岸", "指运港", "运抵国")

_OCR_ENGINE = None


def _get_ocr():
    """懒加载 RapidOCR 引擎（首次调用加载 onnx 模型约 2-3 秒，后续复用）。"""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE


def _ocr_page_to_words(page, dpi=200):
    """扫描件兜底：把 PDF 页渲染成图 → 纠正 90° 旋转 → RapidOCR → 转成 words。

    返回与 pdfplumber `extract_words` 同结构的列表（text/x0/x1/top/bottom，单位 point），
    供下方坐标/标签提取逻辑复用。OCR 失败或空结果返回 []。
    """
    try:
        import numpy as np
        img = page.to_image(resolution=dpi).original
        arr = np.asarray(img)
        # 报关单为横向（842×595）；扫描件若为纵向则逆时针转 90° 还原
        if arr.shape[1] < arr.shape[0]:
            arr = np.rot90(arr, k=1)
        res, _ = _get_ocr()(arr)
    except Exception:
        return []
    words = []
    scale = 72.0 / dpi
    for box, text, _score in (res or []):
        text = (text or "").strip()
        if not text:
            continue
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        words.append({
            "text": text,
            "x0": min(xs) * scale,
            "x1": max(xs) * scale,
            "top": min(ys) * scale,
            "bottom": max(ys) * scale,
        })
    return words


def extract_customs_data(pdf_path, customer=DEFAULT_CUSTOMER):
    """从报关底单PDF提取字段，返回 dict。

    适配「横向 842x595、有文字层」的标准出口货物报关单（铁路/海运/公路）。
    扫描件（无文字层）返回空字段 dict（调用方据此提示人工填写）。
    """
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        # 一票底单可能拆成多份 PDF 合并，且前置可能混入「委托报关协议」等非底单页。
        # 只读第一页会导致第一页不是报关单时字段全空（如「委托报关协议」排最前）。
        # 改为遍历各页，选第一页含报关单关键标签的页作为提取源；找不到再回退第一页（扫描件走 OCR）。
        words = None
        for page in pdf.pages:
            w = page.extract_words(x_tolerance=2, y_tolerance=2)
            if not w:
                continue
            joined = " ".join(x["text"] for x in w)
            if any(lbl in joined for lbl in _BL_KEY_LABELS):
                words = w
                break
        if words is None:
            # 回退：所有页均无文本层报关单标签 → 用第一页（扫描件兜底 OCR，或空底单）
            page = pdf.pages[0]
            words = page.extract_words(x_tolerance=2, y_tolerance=2)
            if not words:
                # 扫描件兜底：无文字层 → OCR 识别（渲染 + 纠偏 + RapidOCR）
                words = _ocr_page_to_words(page)
    if not words:
        return _empty_record(customer)
    lines = _group_by_line(words)  # { y: [words 按 x0 排序] }
    ys = sorted(lines)

    def _row_text(ws):
        return " ".join(w["text"] for w in ws)

    def _value_row(label, dy_min=4, dy_max=20, merge_gap=4):
        """定位含 label 的标签行，返回其下方值行的 words。

        值行可能被 pdfplumber 拆成相邻两行（如公司名 y=101 / 关区名 y=100，基线差 1pt），
        若只取第一行，_leftmost(x0<250) 会漏掉左侧公司名（取到右侧关区名 x0>250 被丢弃）。
        故把 dy 范围内、与首个值行 y 间距 ≤ merge_gap 的相邻行合并返回。
        """
        for y in ys:
            if label in _row_text(lines[y]):
                merged = []
                first_y = None
                for y2 in ys:
                    if dy_min < y2 - y <= dy_max:
                        if first_y is None:
                            first_y = y2
                        elif y2 - first_y > merge_gap:
                            break
                        merged.extend(lines[y2])
                return merged
        return []

    def _leftmost(ws, x_cut=250):
        """返回 x0 < x_cut 的文本（左栏值，通常为发货人/收货人）。"""
        return " ".join(w["text"] for w in ws if w["x0"] < x_cut).strip()

    def _value_below_label(label, x_tol=12, dy_min=4, dy_max=20, gap=15):
        """定位 label 标签，取其紧邻下方 dy 范围内、同 x 列的字段值（可跨多个 word）。

        与 _value_by_label 不同：值行必须落在标签下方 dy_min..dy_max 内，
        避免「运输工具名称及航次号」列值为空/@ 时误抓到下方「征免性质」标签（x0 相同但更远）。
        值可能含空格被拆成多个 word（如「HAWK I/15E」），按 x 连续（gap ≤ 15pt）拼接，
        遇到大 gap（下一字段列，如「提运单号」）即停。
        """
        for y in ys:
            label_x = None
            for w in lines[y]:
                if label in w["text"]:
                    label_x = w["x0"]
                    break
            if label_x is None:
                continue
            for y2 in ys:
                if dy_min < y2 - y <= dy_max:
                    parts = []
                    prev_x1 = None
                    started = False
                    for w in lines[y2]:
                        if w["text"] == "@":
                            continue
                        if not started:
                            # 值起点必须紧贴标签列（x0 相近），否则是本列为空、右侧是别的字段列，
                            # 直接误收「提运单号」等值（铁路底单运输工具列常为空）
                            if abs(w["x0"] - label_x) <= x_tol:
                                started = True
                                parts.append(w["text"])
                                prev_x1 = w["x1"]
                            continue
                        if w["x0"] - prev_x1 > gap:
                            break  # 遇到大 gap：下一字段列，停止
                        parts.append(w["text"])
                        prev_x1 = w["x1"]
                    if parts:
                        return " ".join(parts)
            return ""
        return ""

    data = {}

    # 1. 境内发货人 shipper（值行左侧中文公司名）
    row = _value_row("境内发货人")
    data["shipper"] = _leftmost(row)

    # 2. 境外收货人 consignee（值行左侧英文）
    crow = _value_row("境外收货人")
    data["consignee"] = _leftmost(crow)

    # 3. 提单号 —— 直接定位「提运单号」标签，取同列下方值（如 CHN3541874P8 / G2605115309）
    data["提单号"] = _value_by_label(lines, "提运单号")

    # 4. 运输工具（船名航次）—— 优先「运输工具名称及航次号」标签同列下方值（报关单标准栏位，卡航填车架号如 /91440112M0J494），
    #    过滤占位符 @；取不到回退备注「运输工具名称：XXX」（海运船名航次）
    data["船名航次"] = _value_below_label("运输工具名称及航次号").lstrip("/")
    if not data["船名航次"]:
        data["船名航次"] = _find_captured(lines, r"运输工具(?:名称)?[：:]\s*([^\s;；]+)")
    if not data["船名航次"]:
        vessel = " ".join(w["text"] for w in crow if 380 <= w["x0"] < 500 and w["text"] != "@").strip()
        data["船名航次"] = vessel

    # 5. 集装箱号（柜号）：备注「集装箱标箱数及号码」上下文（避免全文本搜索误匹配提单号）
    data["柜号"] = _container_numbers(lines)

    # 6. 件数 / 毛重 —— 按标签同列下方取值（文本层与 OCR 通用，避免 OCR 各列值错行）
    data["箱数"] = _value_by_label(lines, "件数")
    data["总重量"] = _value_by_label(lines, "毛重")

    # 7. 目的港 / 起运地 / 运抵国 —— 按标签同列下方取值（文本层与 OCR 通用，避免靠行内位置猜列）
    data["起运地"] = _value_by_label(lines, "离境口岸")   # 离境口岸 = 起运地（海运票用）
    data["目的港"] = _value_by_label(lines, "指运港")     # 指运港 = 目的港
    data["运抵国"] = _value_by_label(lines, "运抵国")     # 运抵国（目的国），用于箱数/国家一致性校验

    # 7b. 运输方式 + 关区名 —— 用于起运港分流：海运(水路运输)取离境口岸，非海运取「海关编号」后的关区名
    full_text = " ".join(w["text"] for w in words)
    m = re.search(r"(水路|铁路|公路|航空)运输", full_text)
    data["运输方式"] = m.group(0) if m else ""
    data["关区名"] = _customs_office_from_text(full_text)

    # 8. 起运日期（申报日期，取不到回退出口日期；8 位日期转 2026-07-13 格式）
    d = _value_by_label(lines, "申报日期") or _value_by_label(lines, "出口日期")
    data["起运日期"] = f"{d[:4]}-{d[4:6]}-{d[6:8]}" if re.fullmatch(r"\d{8}", d or "") else ""

    # 9. 品名（明细行「序号 + 10 位商品编号 + 中文名」，去重后拼接）→ 翻译成英文
    data["品名"] = _translate_products(_find_product_names(lines))
    data["总体积"] = ""  # 报关单无体积，由箱货清单提供（见 parse_packing_list）

    # 清理：去括号内容（shipper/consignee 之后会被固定英文覆盖，这里仅兜底）
    for k in ["shipper", "consignee"]:
        data[k] = re.sub(r"[\(（][^)）]*[\)）]", "", data.get(k, "")).strip()
    # 箱数去小数
    if data.get("箱数"):
        try: data["箱数"] = str(int(float(data["箱数"])))
        except: pass
    # 总重量保留 2 位
    if data.get("总重量"):
        try: data["总重量"] = str(round(float(data["总重量"]), 2))
        except: pass

    # 客户固定 shipper/consignee（英文），覆盖提取出的中文/英文名
    cfg = get_customer_config(customer)
    if cfg.get("shipper_bl"):
        data["shipper"] = cfg["shipper_bl"]
    else:
        # 可变发货人（星速 Amazon FBA 直发）：境内发货人中文名 → 拼音大写
        data["shipper"] = _cn_to_pinyin_upper(data.get("shipper", ""))
    data["consignee"] = cfg["consignee_bl"]

    # 星速：合同协议号 = 基础 FBA（用于匹配订单列表 / 箱货清单），文件名兜底在调用方
    if customer == "星速":
        data["FBA"] = _fba_from_text(full_text)

    # 保函字段
    data["申请单位"] = cfg["shipper_telex"]
    data["收货人(保函)"] = cfg["consignee_telex"]  # 可单独改
    data["目的地"]   = data.get("目的港", "")
    # 运输工具 = 船名航次；非海运（铁路/卡航等）底单该列为「@」占位符 → 留空待人工填班列号/车次
    data["运输工具"] = data.get("船名航次", "")
    _now = datetime.now()
    data["申请日期"] = f"{_now.year} 年 {_now.month} 月 {_now.day} 日"
    data["文件命名"] = data.get("提单号", "提单") or "提单"
    return data


def _empty_record(customer=DEFAULT_CUSTOMER):
    """扫描件 / 无文字层时返回的空字段 dict（shipper/consignee 仍预填客户固定英文）。"""
    cfg = get_customer_config(customer)
    return {
        "shipper": cfg["shipper_bl"], "consignee": cfg["consignee_bl"], "提单号": "", "系统SO": "", "柜号": "", "船名航次": "",
        "起运地": "", "目的港": "", "运抵国": "", "运输方式": "", "关区名": "", "箱数": "", "品名": "", "总重量": "", "总体积": "",
        "起运日期": "", "申请单位": cfg["shipper_telex"], "收货人(保函)": cfg["consignee_telex"], "目的地": "", "运输工具": "",
        "申请日期": datetime.now().strftime("%Y 年 %m 月 %d 日"),
        "文件命名": "提单",
    }


# ---------- 2. 生成「审核用 Excel」 ----------
# 提单模板的字段（列顺序），与 ETTON 提单信息收集表一致
BL_FIELDS = [
    "系统SO", "文件命名", "shipper", "consignee", "提单号", "柜号", "船名航次",
    "起运地", "目的港", "箱数", "品名", "总重量", "总体积", "起运日期",
]
# 中文表头
BL_HEADER = {
    "文件命名": "文件命名", "shipper": "Shipper(发货人)", "consignee": "Consignee(收货人)",
    "提单号": "提单号", "系统SO": "系统SO", "柜号": "柜号", "船名航次": "船名航次",
    "起运地": "起运地", "目的港": "目的港", "箱数": "箱数",
    "品名": "品名", "总重量": "总重量", "总体积": "总体积", "起运日期": "起运日期",
}
# 电放保函字段（审核表末尾附加，供人工核对/编辑）
TELEX_FIELDS = ["申请单位", "运输工具", "目的地", "收货人(保函)", "申请日期"]


def build_review_excel(records, xlsx_path):
    """
    records: list[dict]，每个报关单提取出的字段 dict
    生成审核用Excel：第一行是中文表头，每行一票，含 提单字段 + 保函字段。
    同时保留一个「状态」列，方便人工标记。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "提单审核"

    # 表头
    all_cols = ["序号", "状态"] + [BL_HEADER[f] for f in BL_FIELDS] + TELEX_FIELDS
    ws.append(all_cols)

    # 表头样式
    header_fill = PatternFill("solid", fgColor="1E5EFF")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 40

    # 数据行
    for idx, rec in enumerate(records, start=1):
        row = [idx, "待审核"]
        for f in BL_FIELDS:
            row.append(rec.get(f, ""))
        row.append(rec.get("申请单位", ""))
        row.append(rec.get("运输工具", ""))
        row.append(rec.get("目的地", ""))
        row.append(rec.get("收货人(保函)", rec.get("consignee", "")))  # 保函收货人默认=consignee，可单独改
        row.append(rec.get("申请日期", ""))
        ws.append(row)

    # 列宽 & 边框
    widths = {"A": 6, "B": 10, "C": 22, "D": 22, "E": 16, "F": 16, "G": 18,
              "H": 12, "I": 12, "J": 8, "K": 30, "L": 10, "M": 10, "N": 13,
              "O": 22, "P": 16, "Q": 14, "R": 22, "S": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font = Font(size=10)

    # 冻结首行 + 自动筛选
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{chr(64+len(all_cols))}1"

    wb.save(xlsx_path)


def parse_review_excel(xlsx_path):
    """读取前端回传的（经人工审核编辑后的）Excel，返回 records list[dict]"""
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = list(rows[0])
    # 找到英文key映射（用 BL_FIELDS + TELEX 顺序）
    records = []
    for raw in rows[1:]:
        if not raw or all(c is None for c in raw):
            continue
        d = {}
        for i, h in enumerate(headers):
            if i < len(raw):
                d[str(h).strip()] = raw[i]
        # 统一成 key（中文表头 -> 英文key）
        rec = {}
        for eng_key, cn in BL_HEADER.items():
            rec[eng_key] = d.get(cn, d.get(eng_key, ""))
        rec["申请单位"] = d.get("申请单位", rec.get("shipper", ""))
        rec["收货人(保函)"] = d.get("收货人(保函)", rec.get("consignee", ""))
        rec["运输工具"] = d.get("运输工具", rec.get("船名航次", ""))
        rec["目的地"]   = d.get("目的地", rec.get("目的港", ""))
        rec["申请日期"] = d.get("申请日期", datetime.now().strftime("%Y 年 %m 月 %d 日"))
        rec["文件命名"] = rec.get("文件命名") or rec.get("提单号", "提单")
        records.append(rec)
    return records


# ---------- 3. 邮件合并：提单 docx 生成 ----------
_MONTHS_EN = ["", "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
              "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


def _date_to_bl(d):
    """YYYY-MM-DD → DD MMM YYYY（提单日期格式），如 2026-05-14 → 14 MAY 2026。解析失败原样返回。"""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})$", (d or "").strip())
    if not m:
        return (d or "")
    y, mo, dd = m.groups()
    mi = int(mo)
    return f"{dd} {_MONTHS_EN[mi]} {y}" if 1 <= mi <= 12 else (d or "")


_TEMPLATE_CACHE = {}  # 模板路径 -> 已修复模板路径（进程内缓存，幂等）


def _get_deduped_template(template_path):
    """返回「去重重复合并域」后的模板路径（首次调用时修复并缓存）"""
    if template_path in _TEMPLATE_CACHE:
        return _TEMPLATE_CACHE[template_path]
    from clean_template import dedupe_template
    fixed = os.path.join(tempfile.gettempdir(), "etton_bl_模板_已修复.docx")
    dedupe_template(template_path, fixed)
    _TEMPLATE_CACHE[template_path] = fixed
    return fixed


def _fit_field_font(text, box_pt=108.0, default_pt=10.5, min_pt=7.0):
    """估算文本在文本框内的宽度，若超过可用宽度则返回需缩小的字号（0.5pt 步进），否则返回 None（保持模板默认字号）。

    字符宽度系数按 Times New Roman 实测（大写≈0.78em、数字≈0.55em、空格≈0.25em），
    box_pt 取船名航次文本框实测可用宽度（略保守，确保单行不换行）。
    """
    if not text:
        return None
    width_em = 0.0
    for ch in text:
        if ch == " ":
            width_em += 0.25
        elif ch.isdigit():
            width_em += 0.55
        elif ch in "/\\|":
            width_em += 0.35
        elif ch.isupper():
            width_em += 0.78
        elif ch.islower():
            width_em += 0.55
        else:
            width_em += 0.72
    if width_em <= 0:
        return None
    fit = box_pt / width_em
    if fit >= default_pt:
        return None
    return max(round(fit * 2) / 2, min_pt)


def _apply_run_font_size(doc, text, size_pt):
    """给 merge 后文档里值为 text 的 run 设置字号（w:sz / w:szCs），使长文本单行显示。"""
    if size_pt is None or not text:
        return
    from lxml import etree
    from mailmerge import NAMESPACES
    W = NAMESPACES["w"]
    half = str(int(round(size_pt * 2)))
    for part in doc.parts.values():
        for t in part.findall(".//{%s}t" % W):
            if t.text != text:
                continue
            r = t.getparent()
            if r.tag != "{%s}r" % W:
                continue
            rPr = r.find("{%s}rPr" % W)
            if rPr is None:
                rPr = etree.Element("{%s}rPr" % W)
                r.insert(0, rPr)
            for tag in ("{%s}sz" % W, "{%s}szCs" % W):
                el = rPr.find(tag)
                if el is None:
                    el = etree.SubElement(rPr, tag)
                el.set("{%s}val" % W, half)


def fill_bl_docx(template_path, output_path, rec):
    """
    用 docx-mailmerge 填充 ETTON 提单模板的 MERGEFIELD 邮件合并域。
    会先自动修复模板中「重复/嵌套域」（如'目的港'出现多次），保格式不漂移。
    """
    from mailmerge import MailMerge
    import tempfile
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    tpl = _get_deduped_template(template_path)
    doc = MailMerge(tpl)
    mapping = {
        "shipper": str(rec.get("shipper", "")),
        "consignee": str(rec.get("consignee", "")),
        "提单号": str(rec.get("提单号", "")),
        "起运地": str(rec.get("起运地", "")),
        "船名航次": str(rec.get("船名航次", "")),
        "目的港": str(rec.get("目的港", "")),
        "箱数": str(rec.get("箱数", "")),
        "品名": str(rec.get("品名", "")),
        "总重量": str(rec.get("总重量", "")),
        "总体积": str(rec.get("总体积", "")),
        "柜号": str(rec.get("柜号", "")),
        "起运日期": _date_to_bl(rec.get("起运日期", "")),
    }
    doc.merge(**mapping)
    # 船名航次过长时缩小字体，使其在文本框内单行显示（不换行）
    _apply_run_font_size(doc, mapping["船名航次"], _fit_field_font(mapping["船名航次"]))
    doc.write(output_path)


# ---------- 4. 电放保函 docx 生成 ----------
TELEX_TPL = """广州拓锐科技有限公司
电放保函
TELEX RELEASE APPLICATION
TO：{to_company}
FROM：ETTON TECHNOLOGY LOGISTICS (ZHONGSHAN) CO., LTD
提单号 (B/L NO)：{bl_no}
运输工具 (MEANS OF TRANSPORT)：{transport}
目的地 (DESTINATION)：{destination}
收货人 (CONSIGNEE)：{consignee}
以上由本司委托贵司承运的货物，出于我司交易之考虑，特向贵司申请以电放形式于目的港将该货物发放给上述公司，对于由此产生的责任由我司承担。
申请单位：{applicant}
申请单位盖章及申请人签名：
申请日期：{apply_date}"""


def fill_telex_docx(output_path, rec, customer=DEFAULT_CUSTOMER):
    """按参考电放保函模板精确复刻格式（拓锐6月电放保函标准答案）。

    含页眉（客户公司名）+ 黑色分隔线；标题 17/15pt、正文 14pt、声明 12pt；
    中文=宋体/等线、英文数字=Times New Roman；页面边距、header 距离、段落顺序/空行、
    行距、段前间距、左右缩进、全角/半角标点均严格对齐标准答案。
    页眉/TO/FROM 等公司信息取自客户配置（CUSTOMERS），版式暂统一为拓锐版式。"""
    import re
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    cfg = get_customer_config(customer)

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.61)
    sec.bottom_margin = Cm(0)
    sec.left_margin = Cm(3.15)
    sec.right_margin = Cm(2.29)
    sec.header_distance = Cm(1.63)

    L = WD_ALIGN_PARAGRAPH.LEFT
    C = WD_ALIGN_PARAGRAPH.CENTER
    R = WD_ALIGN_PARAGRAPH.RIGHT

    def _run(p, text, ascii_font="Times New Roman", ea_font="宋体", size=14, bold=False):
        r = p.add_run(text)
        r.font.name = ascii_font
        r._element.rPr.rFonts.set(qn("w:eastAsia"), ea_font)
        if size is not None:
            r.font.size = Pt(size)
        if bold:
            r.bold = True
        return r

    def add_mixed(p, text, ascii_font="Times New Roman", ea_font="宋体", size=14, bold=False):
        """按中英文/全角标点拆成多个 run：中文段用 ea_font，英文数字段用 ascii_font。"""
        for seg in re.split(r"([一-鿿　-〿＀-￯]+)", text):
            if seg:
                _run(p, seg, ascii_font, ea_font, size, bold)

    def para(align=None, line=None, before=None, left=None, right=None):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        # 覆盖 python-docx 默认模板的 docDefaults（段后距 10pt、行距 1.15）——否则累积溢出成 2 页
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        if align is not None:
            p.alignment = align
        if line is not None:
            pf.line_spacing = line
        if before is not None:
            pf.space_before = Pt(before)
        if left is not None:
            pf.left_indent = Pt(left)
        if right is not None:
            pf.right_indent = Pt(right)
        return p

    # ---- 页眉：客户公司名 + 黑色分隔线 ----
    hp = sec.header.paragraphs[0]
    hp.alignment = L
    hp.paragraph_format.space_after = Pt(0)
    hp.paragraph_format.line_spacing = 1.0
    _run(hp, cfg["telex_header"], ascii_font="等线", ea_font="等线", size=15.5, bold=True)
    # 右缩进使底边框黑线宽 415.35pt（对齐标准答案 shape 宽度）；space 让黑线下移对齐其位置
    hp.paragraph_format.right_indent = Pt(42.35)
    pPr = hp._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")        # 0.75pt 黑线
    bottom.set(qn("w:space"), "7")     # 黑线与文字间距，下移至标准答案位置（margin-top 73.3pt）
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)

    # ---- 正文（严格对齐参考模板 35 段）----
    # 标题
    p = para(C, line=1.1875, before=3.9)
    add_mixed(p, "电放保函", ascii_font="等线", ea_font="等线", size=17, bold=True)
    p = para(C, before=0, left=57.15, right=55.65)
    add_mixed(p, "TELEX RELEASE APPLICATION", size=15, bold=True)
    # TO / FROM（全角冒号）
    p = para(line=Pt(19.05), before=9.2, left=1.45)
    add_mixed(p, f"TO：{cfg['telex_to']}")
    p = para()
    add_mixed(p, f"FROM ：{cfg['telex_from']}")
    p = para(line=1.7917, before=3.1, left=1.05, right=0.75)
    p = para(line=1.1417)
    # 字段（半角冒号 + 空格）
    p = para(L)
    add_mixed(p, f"提单号 (B/L NO): {rec.get('提单号','')}")
    p = para(L)
    add_mixed(p, " ")
    p = para(L)
    add_mixed(p, f"运输工具 (MEANS OF TRANSPORT): {rec.get('运输工具','')}")
    p = para(L, line=1.0)
    p = para(L)
    add_mixed(p, f"柜号 (CONTAINER NO.): {rec.get('柜号','')}")
    p = para(L, line=1.0)
    p = para(L, line=Pt(15.7))
    add_mixed(p, f"目的地 (DESTINATION): {rec.get('目的地','')}")
    p = para(L, line=1.0)
    p = para(L, line=1.0)
    p = para(L, line=1.0)
    add_mixed(p, f"收货人 (CONSIGNEE): {rec.get('收货人(保函)', rec.get('consignee',''))}")
    p = para(line=1.275)
    p = para(line=1.275)
    p = para(line=1.275)
    # 声明正文（等线 12pt bold，左缩进 8pt，对齐标准答案 Heading3 效果）
    p = para(line=1.0125, left=8)
    add_mixed(p, "以上由本司委托贵司承运 的货物，出于我司交易之考虑，特向贵司申请以电放形式于目的港将该货物发放给上述公司，对于由此产生的责任由我司承担。",
              ascii_font="等线", ea_font="等线", size=12, bold=True)
    p = para(line=1.6417)
    p = para(line=1.6417)
    # 申请单位（全角冒号）
    p = para(line=Pt(19.0), before=4.6, left=5.5)
    add_mixed(p, f"申请单位：{rec.get('申请单位','')}")
    p = para(line=1.0125)
    p = para(line=1.0125)
    p = para(line=1.0125)
    p = para(line=1.0125)
    p = para(line=1.0125)
    p = para(line=1.0167)
    # 盖章签名（等线 14 bold）
    p = para(L, left=8)
    add_mixed(p, "申请单位盖章及申请人签名：", ascii_font="等线", ea_font="等线", size=14, bold=True)
    p = para(L, left=8)
    # 申请日期（右对齐）
    p = para(R, line=0.9458, before=13.8)
    add_mixed(p, f"申请日期：{rec.get('申请日期','')}")
    p = para(R, line=0.9458, before=13.8)
    p = para()
    p = para()

    doc.save(output_path)


# ---------- 5. docx -> PDF 转换 ----------
def _find_soffice():
    """定位 LibreOffice 可执行文件：Windows 原生安装 / Docker Linux 均支持。"""
    for p in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if os.path.exists(p):
            return p
    for cmd in ("soffice", "libreoffice"):
        found = shutil.which(cmd)
        if found:
            return found
    return None


def docx_to_pdf(docx_path, pdf_path):
    """用 LibreOffice headless 转换 docx → PDF。
    找不到 LibreOffice 时静默跳过（仅产 docx），不阻断生成流程。"""
    soffice = _find_soffice()
    if not soffice:
        return
    out_dir = os.path.dirname(pdf_path) or "."
    try:
        subprocess.run([
            soffice, "--headless",
            "--convert-to", "pdf", "--outdir", out_dir, docx_path
        ], check=False, capture_output=True, timeout=120)
    except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
        # LibreOffice 不可用 / 单次转换超时（120s）：跳过 PDF，仅保留 docx，不阻断整批生成
        return
    # soffice 输出同名 pdf，重命名到目标路径
    guessed = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if os.path.exists(guessed) and guessed != pdf_path:
        os.rename(guessed, pdf_path)


# ---------- 6. 一键：审核Excel -> 生成全部 提单+保函 + ZIP ----------
def generate_from_review_xlsx(review_xlsx, bl_template, out_dir, to_pdf=True):
    """
    读取审核后的Excel，逐票生成 提单.docx(+.pdf) + 电放保函.docx(+.pdf)，打包ZIP。
    返回 (zip_path, bl_count, telex_count)
    """
    os.makedirs(out_dir, exist_ok=True)
    records = parse_review_excel(review_xlsx)
    bl_count = 0
    telex_count = 0

    for rec in records:
        name = rec.get("文件命名") or rec.get("提单号") or f"提单_{bl_count+1}"
        safe = re.sub(r'[\\/:*?"<>|]', "_", str(name)).strip() or f"提单_{bl_count+1}"

        # 提单
        bl_docx = os.path.join(out_dir, f"{safe}提单.docx")
        fill_bl_docx(bl_template, bl_docx, rec)
        if to_pdf:
            docx_to_pdf(bl_docx, os.path.join(out_dir, f"{safe}提单.pdf"))
        bl_count += 1

        # 电放保函
        telex_docx = os.path.join(out_dir, f"{safe}电放保函.docx")
        fill_telex_docx(telex_docx, rec)
        if to_pdf:
            docx_to_pdf(telex_docx, os.path.join(out_dir, f"{safe}电放保函.pdf"))
        telex_count += 1

    # 打包 ZIP
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    zip_path = os.path.join(out_dir, f"ETTON提单_电放保函_{stamp}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(os.listdir(out_dir)):
            if f.endswith((".docx", ".pdf")) and not f.startswith("~"):
                zf.write(os.path.join(out_dir, f), arcname=f)
    return zip_path, bl_count, telex_count


# ---------- 7. 批量工作流：拆分底单合并 / 命名 / 国家箱数校验 / 批量生成 ----------
def _declaration_kind(pdf_path):
    """读首页表头判断报关资料类型（不靠文件名，兼容任意命名）。

    返回：0=报关单  1=放行单  2=委托协议  3=其他（排最后，仍合并不丢弃）。
    扫描件（无文字层）读不到文本 → 返回 3（保守保留，靠提取端 OCR 兜底）。
    """
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text() or ""
            if not text.strip():
                # 扫描件（无文字层）→ OCR 首页再判断，保证拆分扫描件也能正确排序（报关单排最前）
                text = " ".join(w["text"] for w in _ocr_page_to_words(page))
    except Exception:
        return 3
    if "委托报关协议" in text:
        return 2  # 委托协议（报关委托书）
    if "出口货物报关单" in text or "进口货物报关单" in text:
        return 0  # 报关单（标题特征最可靠，优先于「通关无纸化」等备注词）
    if "放行通知书" in text:
        return 1  # 放行单
    if "境内发货人" in text:
        return 0  # 报关单（兜底：含「境内发货人」栏位，放行单用「收发货人」）
    return 3


def merge_declaration_pdfs(pdf_paths, output_path):
    """把一票拆分的报关资料（报关单 + 放行单 + 委托协议）合并成一个 PDF。

    顺序固定：报关单 → 放行单 → 委托协议。不靠文件名「报/放/委托」判断，
    而是读每份 PDF 首页表头识别类型后排序；同类型按文件名稳定排序。
    三类都合并（委托协议也是报关资料的一部分），只排序不丢弃。
    """
    from pypdf import PdfWriter
    ordered = sorted(pdf_paths, key=lambda p: (_declaration_kind(p), os.path.basename(p)))
    writer = PdfWriter()
    for p in ordered:
        writer.append(p)
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


def derive_output_name(folder_name, kind):
    """文件夹名 -> 输出文件名：把「易通报关资料」替换成 kind（提单/保函/底单），其余严格照抄。

    例：'6月第1周-(易通报关资料）-BG20260605002  10票  英国铁路不包税卡派 54箱  是4'
      -> '6月第1周-(提单）-BG20260605002  10票  英国铁路不包税卡派 54箱  是4'
    """
    return str(folder_name).replace("易通报关资料", kind)


_COUNTRIES = [
    "英国", "美国", "加拿大", "德国", "法国", "西班牙", "意大利", "荷兰",
    "波兰", "澳大利亚", "日本", "韩国", "俄罗斯", "墨西哥", "巴西", "印度",
    "阿联酋", "沙特", "新加坡", "马来西亚", "泰国", "越南",
]


def parse_folder_country(folder_name):
    """从文件夹名解析目的国（如 '英国快铁不包税-卡派' -> '英国'）。找不到返回 None。"""
    for c in _COUNTRIES:
        if c in str(folder_name):
            return c
    return None


def parse_folder_box_count(folder_name):
    """从文件夹名解析总箱数（如 '54箱' -> 54）。找不到返回 None。"""
    m = re.search(r"(\d+)\s*箱", str(folder_name))
    return int(m.group(1)) if m else None


def _to_int(v):
    try:
        if v is None or str(v).strip() == "":
            return None
        return int(float(v))
    except (ValueError, TypeError):
        return None


def _to_float(v):
    try:
        if v is None or str(v).strip() == "":
            return None
        return float(v)
    except (ValueError, TypeError):
        return None


def validate_ticket(folder, rec, packing):
    """对比 报关底单 / 箱货清单 / 文件夹名 的「国家」和「箱数」，返回不一致提醒列表。

    - 底单国家 = rec['运抵国']，箱数 = rec['箱数']
    - 箱货清单国家 = packing['country']，箱数 = packing['total_boxes']
    - 文件夹名国家 = parse_folder_country(folder)，箱数 = parse_folder_box_count(folder)
    三者只要有值且互相不同，就提醒「原底单/箱货清单有问题」。
    """
    warns = []

    fbox = parse_folder_box_count(folder)
    cbox = _to_int(rec.get("箱数"))
    pbox = _to_int(packing.get("total_boxes")) if packing else None
    boxes = {}
    if fbox is not None:
        boxes["文件夹名"] = fbox
    if cbox is not None:
        boxes["报关底单"] = cbox
    if pbox is not None:
        boxes["箱货清单"] = pbox
    if len({v for v in boxes.values()}) > 1:
        detail = "、".join(f"{k}={v}" for k, v in boxes.items())
        warns.append(f"箱数不一致（{detail}），请核对原底单/箱货清单")

    fcountry = parse_folder_country(folder)
    ccountry = (rec.get("运抵国") or "").strip() or None
    pcountry = (packing.get("country") or "").strip() if packing else None
    countries = {}
    if fcountry:
        countries["文件夹名"] = fcountry
    if ccountry:
        countries["报关底单"] = ccountry
    if pcountry:
        countries["箱货清单"] = pcountry
    if len({c for c in countries.values() if c}) > 1:
        detail = "、".join(f"{k}={v}" for k, v in countries.items())
        warns.append(f"国家不一致（{detail}），请核对原底单/箱货清单")

    return warns


def generate_batch(tickets, bl_template, out_dir, customer=DEFAULT_CUSTOMER, root_folder=""):
    """批量生成：每票 提单.pdf + 电放保函.docx + 底单.pdf + 保函预览PDF，打包 ZIP（仅前三份）。

    tickets: list[dict]，每项含
        folder     文件夹名（用于命名）
        record     审核后字段 dict
        merged_pdf 该票合并后的报关底单 PDF 路径
    customer: 客户标识（决定保函页眉/TO/FROM 等；提单暂共用模板）。
    返回 (zip_path, results)；results 每项含 bl_pdf / telex_docx / telex_preview_pdf / ddan_pdf 及命名。
    """
    os.makedirs(out_dir, exist_ok=True)
    preview_dir = os.path.join(out_dir, "preview")
    os.makedirs(preview_dir, exist_ok=True)

    no_telex = bool(get_customer_config(customer).get("no_telex"))
    is_xs = (customer == "星速")

    results = []
    for i, t in enumerate(tickets):
        rec = t["record"]
        folder = t.get("folder") or rec.get("提单号") or f"提单_{i+1}"

        # 星速命名对齐历史提单：FBA15XXX-公司-件数-提单.pdf / -底单.pdf
        if is_xs:
            bl_name = f"{folder}-提单"
            telex_name = ""
            ddan_name = f"{folder}-底单"
        else:
            bl_name = derive_output_name(folder, "提单")
            telex_name = derive_output_name(folder, "电放保函")
            ddan_name = derive_output_name(folder, "底单")
        safe_bl = re.sub(r'[\\/:*?"<>|]', "_", bl_name) or "提单"
        safe_telex = re.sub(r'[\\/:*?"<>|]', "_", telex_name) or "电放保函"
        safe_ddan = re.sub(r'[\\/:*?"<>|]', "_", ddan_name) or "底单"

        # 提单：docx 临时 -> PDF（仅保留 PDF）。所有客户统一走 ETTON 邮件合并模板（星速同样套用拓锐模板）。
        bl_docx_tmp = os.path.join(tempfile.gettempdir(), f"{safe_bl}_{i}.docx")
        bl_pdf = os.path.join(out_dir, f"{safe_bl}.pdf")
        fill_bl_docx(bl_template, bl_docx_tmp, rec)
        docx_to_pdf(bl_docx_tmp, bl_pdf)
        try:
            os.remove(bl_docx_tmp)
        except OSError:
            pass

        # 电放保函：docx（ZIP 用）+ 预览 PDF（不进 ZIP）。星速（Amazon FBA 直发）无需保函，跳过。
        telex_docx = ""
        telex_preview_pdf = ""
        if not no_telex:
            telex_docx = os.path.join(out_dir, f"{safe_telex}.docx")
            fill_telex_docx(telex_docx, rec, customer)
            telex_preview_pdf = os.path.join(preview_dir, f"{safe_telex}.pdf")
            docx_to_pdf(telex_docx, telex_preview_pdf)

        # 底单：合并 PDF 复制改名（星速 ZIP 不含底单，跳过复制）
        ddan_pdf = os.path.join(out_dir, f"{safe_ddan}.pdf")
        merged = t.get("merged_pdf")
        if not is_xs and merged and os.path.exists(merged):
            shutil.copy(merged, ddan_pdf)

        results.append({
            "folder": folder,
            "bl_pdf": bl_pdf,
            "telex_docx": telex_docx,
            "telex_preview_pdf": telex_preview_pdf,
            "ddan_pdf": ddan_pdf,
            "bl_name": safe_bl, "telex_name": safe_telex, "ddan_name": safe_ddan,
        })

    # 打包 ZIP：拓锐按票建子文件夹（底单+提单+保函）；星速所有提单平铺根目录（不建子文件夹、不含底单）
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    safe_root = re.sub(r'[\\/:*?"<>|]', "_", (root_folder or "").strip())
    zip_name = f"{safe_root}系统制作文件.zip" if safe_root else (f"ETTON提单_{stamp}.zip" if is_xs else f"ETTON提单_电放保函_{stamp}.zip")
    zip_path = os.path.join(out_dir, zip_name)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        if is_xs:
            # 星速：所有提单平铺根目录（命名 = 底单名前缀 + 「-提单」，见 bl_name）
            for r in results:
                src = r["bl_pdf"]
                if src and os.path.isfile(src):
                    zf.write(src, arcname=os.path.basename(src))
        else:
            for j, r in enumerate(results):
                safe_folder = (re.sub(r'[\\/:*?"<>|]', "_", str(r.get("folder") or "")).strip()
                               or f"票_{j + 1}")
                # 顺序对齐用户预期：底单、提单、保函（底单缺失则跳过，不阻断）
                for src in (r["ddan_pdf"], r["bl_pdf"], r["telex_docx"]):
                    if src and os.path.isfile(src):
                        zf.write(src, arcname=f"{safe_folder}/{os.path.basename(src)}")
    return zip_path, results


# ---------- 内部工具 ----------
def _group_by_line(words):
    lines = {}
    for w in words:
        y = round(w["top"], 0)
        lines.setdefault(y, []).append(w)
    for y in lines:
        lines[y].sort(key=lambda w: w["x0"])
    return lines

def _get_text_at(lines, y_target, x_min=None, x_max=None, y_tol=3):
    for y, words in lines.items():
        if abs(y - y_target) <= y_tol:
            texts = [w["text"] for w in words
                     if (x_min is None or w["x0"] >= x_min)
                     and (x_max is None or w["x0"] <= x_max)]
            return " ".join(texts).strip()
    return ""

def _value_by_label(lines, label, x_tol=12):
    """定位 label 标签，返回其下方同 x 列（x0 相近）的第一个非占位值。

    报关单字段值通常在标签正下方、与标签共用同一列（如「提运单号」标签 x≈517，
    值在下方同列，海运/铁路布局通用）。找不到返回空串。
    """
    for y in sorted(lines):
        label_x = None
        for w in lines[y]:
            if label in w["text"]:
                label_x = w["x0"]
                break
        if label_x is None:
            continue
        for y2 in sorted(lines):
            if y2 <= y:
                continue
            for w in lines[y2]:
                if abs(w["x0"] - label_x) <= x_tol and w["text"] not in ("@", label):
                    return w["text"]
        return ""
    return ""

def _find_in_notes(lines, pattern):
    for y, words in lines.items():
        if 235 <= y <= 250:
            full = " ".join(w["text"] for w in words)
            m = re.search(pattern, full)
            if m: return m.group(0)
    return ""

def _find_any(lines, pattern):
    for words in lines.values():
        full = " ".join(w["text"] for w in words)
        m = re.search(pattern, full)
        if m: return m.group(0)
    return ""

def _find_captured(lines, pattern):
    """全文本搜索，返回第一个捕获组（找不到返回空串）。"""
    for words in lines.values():
        full = " ".join(w["text"] for w in words)
        m = re.search(pattern, full)
        if m: return m.group(1)
    return ""


def _container_numbers(lines):
    """从备注「集装箱标箱数及号码」上下文提取柜号（4 字母 + 7 数字），多柜用分号拼接。

    不能全文本 [A-Z]{4}\\d{7} 搜索——海运提单号（如 ZIMUNGB1391012S）也含此格式会被误匹配；
    柜号固定出现在备注行「集装箱标箱数及号码：N;XXXX1234567;」里（海运/铁路底单格式一致）。
    """
    for words in lines.values():
        full = " ".join(w["text"] for w in words)
        m = re.search(r"集装箱标箱数及号码\s*[：:]?\s*\d+\s*[;；]", full)
        if not m:
            continue
        nums = re.findall(r"[A-Z]{4}\d{7}", full[m.end():])
        if nums:
            seen = []
            for n in nums:
                if n not in seen:
                    seen.append(n)
            return ";".join(seen)
    return ""

def _find_product_names(lines):
    """提取明细行的中文品名，去重后拼接。

    兼容两种形态：
      - 文本层「序号 + 10 位商品编号 + 中文名」（如 `1 9405429000 水波纹灯`）
      - OCR「10 位编号紧贴中文名」（如 `9405429000水波纹灯`）
    """
    names = []
    for y, words in lines.items():
        full = " ".join(w["text"] for w in words)
        m = re.match(r"^\d+\s+\d{10}\s*([一-鿿]+)", full)
        if not m:
            m = re.search(r"(?<!\d)\d{10}\s*([一-鿿][一-鿿]*)\s*$", full)
        if m:
            name = m.group(1)
            if name not in names:
                names.append(name)
    return "、".join(names)


def _strip_version(name):
    """去掉品名末尾的版本号（如 'Projector 3.0'→'Projector'、'Projector2.0'→'Projector'），其余大小写原样保留。"""
    if not name:
        return ""
    return re.sub(r"\s*\d+(?:\.\d+)+\s*$", "", name).strip()


def _translate_products(cn_names_str):
    """中文品名 -> 英文品名（PRODUCT_EN_MAP），未命中保留原文；去版本号 + 保持大小写，换行拼接（提单每行一个品名）。"""
    if not cn_names_str:
        return ""
    parts = re.split(r"[、,，;；/]", cn_names_str)
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        en = PRODUCT_EN_MAP.get(p, p)
        out.append(_strip_version(en).upper())
    return "\n".join(out)


def parse_packing_list(xlsx_path):
    """解析箱货清单 xlsx，返回 { products_en, total_volume, total_weight }。

    - products_en: 去重后的英文品名（大写，按出现顺序），用于覆盖提单「品名」
    - total_volume: 该票总体积(CBM)——有「体积」列则直接求和；无则回退「长(CM)×宽(CM)×高(CM)×总箱数/1e6」
    - total_weight: 各工作号「货箱重量」列求和 = 该票总毛重(kg)
    - fba_ids: 去重后的 FBA ID（用于去物流追踪表匹配 ETD 开船日）
    """
    wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    en_col = cn_col = vol_col = wt_col = None
    len_col = wid_col = hei_col = box_col = None
    fba_col = channel_col = country_col = so_col = None
    header_idx = None
    for i, row in enumerate(rows[:12]):
        if not row:
            continue
        for j, v in enumerate(row):
            s = str(v).strip() if v is not None else ""
            if s == "英文品名":
                en_col = j
                header_idx = i
            elif s == "中文品名":
                cn_col = j
            elif s == "体积":
                vol_col = j
            elif s == "货箱重量":
                wt_col = j
            elif s == "长(CM)":
                len_col = j
            elif s == "宽(CM)":
                wid_col = j
            elif s == "高(CM)":
                hei_col = j
            elif s == "总箱数(CTN)":
                box_col = j
            elif s == "FBA ID":
                fba_col = j
            elif s == "客户渠道":
                channel_col = j
            elif s in ("国家", "目的国"):
                country_col = j
            elif s == "系统SO":
                so_col = j

    products_en = []
    fba_ids = []
    so_numbers = []
    total_vol = 0.0
    total_wt = 0.0
    total_boxes = 0
    channel = None
    country = None
    for i, row in enumerate(rows):
        if not row:
            continue
        # 跳过表头行及其之前的行，避免把「英文品名」表头本身当成品名
        if header_idx is not None and i <= header_idx:
            continue
        if fba_col is not None and fba_col < len(row) and row[fba_col]:
            fid = str(row[fba_col]).strip()
            if fid and fid not in fba_ids:
                fba_ids.append(fid)
        if channel_col is not None and channel_col < len(row) and row[channel_col]:
            ch = str(row[channel_col]).strip()
            if ch and channel is None:
                channel = ch
        if country_col is not None and country_col < len(row) and row[country_col]:
            c = str(row[country_col]).strip()
            if c and country is None:
                country = c
        if so_col is not None and so_col < len(row) and row[so_col]:
            so = str(row[so_col]).strip()
            if so and so not in so_numbers:
                so_numbers.append(so)
        if en_col is not None and en_col < len(row):
            en = row[en_col]
            if en:
                e = _strip_version(str(en)).upper()
                if e and e not in products_en:
                    products_en.append(e)
        if vol_col is not None and vol_col < len(row):
            v = row[vol_col]
            if isinstance(v, (int, float)):
                total_vol += float(v)
        elif None not in (len_col, wid_col, hei_col, box_col) and max(len_col, wid_col, hei_col, box_col) < len(row):
            # 无「体积」列时回退：长(cm)*宽(cm)*高(cm)*总箱数 / 1e6 = CBM
            L, W, H, B = row[len_col], row[wid_col], row[hei_col], row[box_col]
            if all(isinstance(x, (int, float)) for x in (L, W, H, B)):
                total_vol += float(L) * float(W) * float(H) * float(B) / 1e6
        if wt_col is not None and wt_col < len(row):
            v = row[wt_col]
            if isinstance(v, (int, float)):
                total_wt += float(v)
        if box_col is not None and box_col < len(row):
            b = row[box_col]
            if isinstance(b, (int, float)):
                total_boxes += int(b)
    return {
        "products_en": products_en,
        "total_volume": round(total_vol, 4),
        "total_weight": round(total_wt, 2),
        "total_boxes": total_boxes,
        "fba_ids": fba_ids,
        "so_numbers": so_numbers,
        "channel": channel,
        "country": country,
    }


# ---------- 周汇总箱货清单：按工作号分组 + 匹配报关底单 ----------
def parse_packing_list_weekly(xlsx_path):
    """按「工作号」分组解析周汇总箱货清单（整周多票合并在一份 xlsx），返回分组列表。

    每项：{ 工作号, country, channel, boxes, weight, volume, products_en, fba_ids }
    一张报关底单可能对应多个工作号（合并报关），匹配阶段由 match_weekly_packing 组合。
    无「工作号」列时返回空列表（调用方回退到单票 parse_packing_list）。
    """
    wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    en_col = vol_col = wt_col = None
    len_col = wid_col = hei_col = box_col = None
    fba_col = channel_col = country_col = work_col = so_col = None
    header_idx = None
    for i, row in enumerate(rows[:12]):
        if not row:
            continue
        for j, v in enumerate(row):
            s = str(v).strip() if v is not None else ""
            if s == "英文品名":
                en_col = j
                header_idx = i
            elif s == "体积":
                vol_col = j
            elif s == "货箱重量":
                wt_col = j
            elif s == "长(CM)":
                len_col = j
            elif s == "宽(CM)":
                wid_col = j
            elif s == "高(CM)":
                hei_col = j
            elif s == "总箱数(CTN)":
                box_col = j
            elif s == "FBA ID":
                fba_col = j
            elif s == "客户渠道":
                channel_col = j
            elif s in ("国家", "目的国"):
                country_col = j
            elif s == "工作号":
                work_col = j
            elif s == "系统SO":
                so_col = j

    if work_col is None:
        return []

    groups = {}
    for i, row in enumerate(rows):
        if not row:
            continue
        if header_idx is not None and i <= header_idx:
            continue
        wid = row[work_col] if work_col < len(row) else None
        if not wid:
            continue
        wid = str(wid).strip()
        g = groups.setdefault(wid, {
            "工作号": wid, "country": None, "channel": None, "so": None,
            "boxes": 0, "weight": 0.0, "volume": 0.0,
            "products_en": [], "fba_ids": [],
        })
        if so_col is not None and so_col < len(row) and row[so_col] and g["so"] is None:
            g["so"] = str(row[so_col]).strip()
        if country_col is not None and country_col < len(row) and row[country_col] and g["country"] is None:
            g["country"] = str(row[country_col]).strip()
        if channel_col is not None and channel_col < len(row) and row[channel_col] and g["channel"] is None:
            g["channel"] = str(row[channel_col]).strip()
        if fba_col is not None and fba_col < len(row) and row[fba_col]:
            fid = str(row[fba_col]).strip()
            if fid and fid not in g["fba_ids"]:
                g["fba_ids"].append(fid)
        if en_col is not None and en_col < len(row) and row[en_col]:
            e = _strip_version(str(row[en_col])).upper()
            if e and e not in g["products_en"]:
                g["products_en"].append(e)
        if vol_col is not None and vol_col < len(row):
            v = row[vol_col]
            if isinstance(v, (int, float)):
                g["volume"] += float(v)
        elif None not in (len_col, wid_col, hei_col, box_col) and max(len_col, wid_col, hei_col, box_col) < len(row):
            L, W, H, B = row[len_col], row[wid_col], row[hei_col], row[box_col]
            if all(isinstance(x, (int, float)) for x in (L, W, H, B)):
                g["volume"] += float(L) * float(W) * float(H) * float(B) / 1e6
        if wt_col is not None and wt_col < len(row):
            v = row[wt_col]
            if isinstance(v, (int, float)):
                g["weight"] += float(v)
        if box_col is not None and box_col < len(row):
            b = row[box_col]
            if isinstance(b, (int, float)):
                g["boxes"] += int(b)

    return [g for g in groups.values()]


def match_weekly_packing(record, workgroups):
    """把一张报关底单 record 匹配到周汇总箱货清单的若干工作号，返回合并后的箱货清单 dict。

    匹配键：运抵国 + 件数(箱数) + 毛重（容差 1.0kg）。返回 None 表示匹配失败
    （扫描件无文字层 / 字段缺失 / 凑不出箱数+重量组合）。
    返回结构兼容 parse_packing_list（products_en/total_volume/total_weight/total_boxes/fba_ids/channel/country）。
    """
    ctry = (record.get("运抵国") or "").strip()
    boxes = _to_int(record.get("箱数"))
    weight = _to_float(record.get("总重量"))
    if not ctry or boxes is None or weight is None:
        return None

    cands = [g for g in workgroups if (g.get("country") or "").strip() == ctry]
    if not cands:
        return None
    # 先试单工作号直接命中（常见一票 = 一工作号）
    for g in cands:
        if g["boxes"] == boxes and abs(g["weight"] - weight) < 1.0:
            return _merge_workgroups([g])
    # 否则回溯凑组合（一票 = 多工作号合并报关，如 6月第4周 BG20260625002 = 2 个工作号）
    chosen = _subset_sum(cands, boxes, weight, 0, [], 0, 0.0)
    if not chosen:
        return None
    return _merge_workgroups(chosen)


def _subset_sum(cands, target_boxes, target_wt, idx, chosen, cur_boxes, cur_wt):
    """回溯找「箱数之和==target_boxes 且 重量之和≈target_wt」的工作号子集。"""
    if cur_boxes == target_boxes and abs(cur_wt - target_wt) < 1.0:
        return list(chosen)
    if cur_boxes >= target_boxes or idx >= len(cands):
        return None
    r = _subset_sum(cands, target_boxes, target_wt, idx + 1, chosen, cur_boxes, cur_wt)
    if r is not None:
        return r
    g = cands[idx]
    chosen.append(g)
    r = _subset_sum(cands, target_boxes, target_wt, idx + 1, chosen, cur_boxes + g["boxes"], cur_wt + g["weight"])
    chosen.pop()
    return r


def _merge_workgroups(groups):
    """合并若干工作号组为单票箱货清单 dict（结构对齐 parse_packing_list 返回）。"""
    products, fba_ids, so_numbers = [], [], []
    total_vol = total_wt = 0.0
    total_boxes = 0
    channel = country = None
    for g in groups:
        if channel is None:
            channel = g.get("channel")
        if country is None:
            country = g.get("country")
        for e in g.get("products_en", []):
            if e not in products:
                products.append(e)
        for f in g.get("fba_ids", []):
            if f not in fba_ids:
                fba_ids.append(f)
        if g.get("so") and g["so"] not in so_numbers:
            so_numbers.append(g["so"])
        total_vol += g.get("volume", 0.0)
        total_wt += g.get("weight", 0.0)
        total_boxes += g.get("boxes", 0)
    return {
        "products_en": products,
        "total_volume": round(total_vol, 4),
        "total_weight": round(total_wt, 2),
        "total_boxes": total_boxes,
        "fba_ids": fba_ids,
        "so_numbers": so_numbers,
        "channel": channel,
        "country": country,
    }


# ---------- 物流追踪表：FBA ID -> ETD 开船日 ----------
def _excel_date_str(v):
    """Excel 日期值 -> 'YYYY-MM-DD' 字符串（数字序列号 / datetime / 字符串均兼容）。"""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime("%Y-%m-%d")
    s = str(v).strip()
    # 字符串日期可能带时间/小数（如 2026-07-03 00:00:00 / 2026-07-03 00:00:00.0）→ 只取日期
    m = re.match(r"(\d{4}-\d{2}-\d{2})", s)
    return m.group(1) if m else s


def parse_tracking_list(xlsx_path):
    """解析物流追踪表 xlsx，返回 { FBA ID: {etd, eta, vessel} }。

    物流追踪表多个 sheet（2023/2024/2025/最新物流动态/10月），表头位置不一：
    - 「Shipment ID（FBA号）」定位 FBA 列
    - 「ETD」定位开船日列（etd，起运日期）
    - 「ETA」定位到港日列（eta，算到港时间/申请日期）
    - 「船名航次」/「航次」/「班列」定位船名航次列（vessel，铁路为班列号）
    遍历所有 sheet，按表头文字定位列，合并成 FBA ID -> 信息索引（后 sheet 字段级覆盖先 sheet）。
    """
    wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    index = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        sid_col = etd_col = eta_col = vessel_col = None
        header_row = None
        for i, row in enumerate(rows[:5]):
            if not row:
                continue
            for j, v in enumerate(row):
                s = str(v).strip() if v is not None else ""
                if "Shipment ID" in s or "FBA号" in s:
                    sid_col = j
                if "ETA" in s:
                    eta_col = j
                if "ETD" in s:
                    etd_col = j
                if "船名" in s or "班列" in s or "航次" in s:
                    vessel_col = j
            if sid_col is not None and (etd_col is not None or eta_col is not None):
                header_row = i
                break
        if header_row is None:
            continue
        for row in rows[header_row + 1:]:
            if not row:
                continue
            sid = str(row[sid_col]).strip() if sid_col < len(row) and row[sid_col] else ""
            if not sid:
                continue
            info = index.setdefault(sid, {})
            if etd_col is not None and etd_col < len(row) and row[etd_col]:
                info["etd"] = _excel_date_str(row[etd_col])
            if eta_col is not None and eta_col < len(row) and row[eta_col]:
                info["eta"] = _excel_date_str(row[eta_col])
            if vessel_col is not None and vessel_col < len(row) and row[vessel_col]:
                info["vessel"] = str(row[vessel_col]).strip()
    wb.close()
    return index


def _arrival_date(etd, eta, express=False):
    """到港时间 = ETA 提前两天；若 ETA - ETD 不足 2 天，直接取 ETA。
    缺 ETA 时（如铁路票 ETA 未填）退而求其次取 ETD + 10 天；两者皆缺返回空串。
    快渠道（express=True，空运/国际快递）缺 ETA 不做 ETD+10 兜底，返回空串。"""
    eta = (eta or "").strip()
    etd = (etd or "").strip()
    if not eta:
        # 无 ETA：快渠道（空运/快递）直接返回空；慢渠道用 ETD + 10 天兜底
        if express or not etd:
            return ""
        try:
            return (datetime.strptime(etd, "%Y-%m-%d") + timedelta(days=10)).strftime("%Y-%m-%d")
        except ValueError:
            return ""
    try:
        eta_d = datetime.strptime(eta, "%Y-%m-%d")
    except ValueError:
        return eta
    if etd:
        try:
            etd_d = datetime.strptime(etd, "%Y-%m-%d")
            if (eta_d - etd_d).days < 2:
                return eta
        except ValueError:
            pass
    return (eta_d - timedelta(days=2)).strftime("%Y-%m-%d")


def _date_cn(s):
    """'YYYY-MM-DD' → 'Y 年 M 月 D 日'（对齐电放保函「申请日期」格式，不补零）。"""
    try:
        d = datetime.strptime((s or "").strip(), "%Y-%m-%d")
        return f"{d.year} 年 {d.month} 月 {d.day} 日"
    except ValueError:
        return (s or "").strip()


if __name__ == "__main__":
    # 自测：单文件流程
    pdf = "/data/inputs/QZKT26060001-24件报关底单.pdf"
    tpl = "/data/inputs/提单模板.docx"
    base = tempfile.mkdtemp()
    rec = extract_customs_data(pdf)
    print("提取字段:")
    for k, v in rec.items():
        print(f"  {k}: {v}")
    xlsx = os.path.join(base, "审核表.xlsx")
    build_review_excel([rec], xlsx)
    print("\n审核Excel已生成:", xlsx)
    zip_path, b, t = generate_from_review_xlsx(xlsx, tpl, base, to_pdf=False)
    print(f"生成完成: {zip_path} (提单{b}, 保函{t})")
