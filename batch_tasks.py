# -*- coding: utf-8 -*-
"""
统一运营管理平台 —— 任务批量创建脚本（BrowserSkill/bsk 驱动）

数据来源：一份 Excel，一行 = 一条任务。行数不设上限，读多少跑多少。

用法（全部可选，不带参数 = 跑完表里所有还没跑过的行）：
  python batch_tasks.py                       # 全量，自动跳过已提交过的（断点续跑）
  python batch_tasks.py --list                # 只看将要处理的数据，不碰浏览器
  python batch_tasks.py --limit 20            # 本次最多跑 20 条就收工
  python batch_tasks.py --from 100 --to 200   # 只跑编号 100..200
  python batch_tasks.py --only 7              # 只跑第 7 条
  python batch_tasks.py --dry-run             # 填完不提交，停在提交前（配合 --only 单条检查）
  python batch_tasks.py --force               # 忽略「已提交」记录，强制重跑
  python batch_tasks.py --excel D:/x.xlsx --sheet 浏览任务 --session ojqo
  python batch_tasks.py --reset-state         # 清空断点记录后退出

配置见 config.json；命令行参数优先级高于配置。
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.parse
import zipfile
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.abspath(__file__))
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

RE_REF = re.compile(r'^(\s*)@(e\d+)\s+([A-Za-z]+)(?:\s+"((?:[^"\\]|\\.)*)")?')
RE_VALUE = re.compile(r'="((?:[^"\\]|\\.)*)"\s*$')


class StepError(RuntimeError):
    pass


# ---------------------------------------------------------------- 配置与日志
def load_cfg(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


class Log:
    def __init__(self, path):
        self.fh = open(path, "a", encoding="utf-8")

    def __call__(self, msg, level="INFO"):
        line = f"[{dt.datetime.now():%H:%M:%S}] {level:<5} {msg}"
        print(line, flush=True)
        self.fh.write(line + "\n")
        self.fh.flush()


# ---------------------------------------------------------------- 读 Excel
def read_xlsx_rows(path, sheet_name, cols, start_row):
    """返回 {行号: {列字母: 值}}，不依赖第三方库。cols 可为空集合表示全部列。"""
    def col_of(ref):
        return re.match(r"([A-Z]+)", ref).group(1)

    with zipfile.ZipFile(path) as z:
        try:
            sst_root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            sst = ["".join(t.text or "" for t in si.iter(NS + "t")) for si in sst_root.findall(NS + "si")]
        except KeyError:
            sst = []

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid2target = {r.get("Id"): r.get("Target") for r in rels}

        target = None
        for sh in wb.find(NS + "sheets"):
            if sh.get("name") == sheet_name:
                target = rid2target[sh.get(REL_NS + "id")]
                break
        if target is None:
            raise StepError(f"工作表 {sheet_name} 不存在（可用：{[s.get('name') for s in wb.find(NS + 'sheets')]}）")
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")

        sheet = ET.fromstring(z.read(target))
        rows = {}
        for row in sheet.iter(NS + "row"):
            rno = int(row.get("r"))
            if rno < start_row:
                continue
            vals = {}
            for c in row.findall(NS + "c"):
                col = col_of(c.get("r") or "")
                if cols and col not in cols:
                    continue
                t = c.get("t")
                v = c.find(NS + "v")
                if t == "s" and v is not None:
                    val = sst[int(v.text)]
                elif t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter(NS + "t"))
                else:
                    val = v.text if v is not None else ""
                vals[col] = (val or "").strip()
            if any(vals.values()):
                rows[rno] = vals
        return rows


# ---------------------------------------------------------------- 图片托管
def ensure_served(src, image_dir, base_url):
    """把图片复制进本地服务的目录，返回可被页面 fetch 的 URL。"""
    src = os.path.abspath(src)
    if not os.path.isfile(src):
        raise StepError(f"图片不存在：{src}")
    os.makedirs(image_dir, exist_ok=True)
    ext = os.path.splitext(src)[1].lower() or ".png"
    name = hashlib.sha1(src.encode("utf-8")).hexdigest()[:16] + ext
    dst = os.path.join(image_dir, name)
    if not os.path.isfile(dst) or os.path.getsize(dst) != os.path.getsize(src):
        shutil.copyfile(src, dst)
    return base_url.rstrip("/") + "/" + urllib.parse.quote(name)


# ---------------------------------------------------------------- 断点记录
class Ledger:
    """记录已成功提交的任务名，重跑时自动跳过，避免重复提交。"""

    def __init__(self, path, log):
        self.path = path
        self.log = log
        self.data = {"items": {}}
        if os.path.isfile(path):
            try:
                with open(path, encoding="utf-8") as f:
                    self.data = json.load(f)
                self.data.setdefault("items", {})
            except Exception:  # noqa: BLE001
                self.log(f"断点文件损坏，已重建：{path}", "WARN")

    @staticmethod
    def key(name):
        return re.sub(r"\s+", "", name or "")

    def has(self, name):
        return self.key(name) in self.data["items"]

    def add(self, name, meta):
        self.data["items"][self.key(name)] = meta
        self._flush()

    def _flush(self):
        os.makedirs(os.path.dirname(self.path) or ".", exist_ok=True)
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.path)

    def reset(self):
        self.data = {"items": {}}
        self._flush()


# ---------------------------------------------------------------- bsk 封装
class Browser:
    def __init__(self, cfg, log):
        self.cfg = cfg
        self.log = log
        self.session = cfg["session"]
        self.env = dict(os.environ, BSK_AUTO_START="0")

    def run(self, args, timeout=None, check=True):
        timeout = timeout or self.cfg.get("step_timeout", 60)
        cmd = [self.cfg.get("bsk_exe", "bsk")] + args
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=timeout, env=self.env)
        out = (p.stdout or "") + (p.stderr or "")
        if check and p.returncode != 0:
            raise StepError(f"bsk {' '.join(args[:2])} 失败：{out.strip()[:400]}")
        return out

    # --- 同步 / 观察 -------------------------------------------------
    def sync(self):
        out = self.run(["cooperate", "sync", "--session", self.session, "--json"])
        data = json.loads(out[out.find("{"):])
        return data["context_id"], data.get("observation_id")

    def observe(self):
        return self.run(["observe", "--session", self.session])

    # --- 解析 VOM ---------------------------------------------------
    @staticmethod
    def parse(vom):
        items = []
        for line in vom.splitlines():
            m = RE_REF.match(line)
            if not m:
                continue
            indent, ref, kind, name = m.group(1), m.group(2), m.group(3), m.group(4)
            name = (name or "").replace('\\"', '"').replace("\\\\", "\\")
            # 行尾的 ="..." 可能是 placeholder 而不是值，必须先剔除 placeholder
            value = None
            if "[empty]" not in line:
                clean = re.sub(r'placeholder="(?:[^"\\]|\\.)*"', "", line)
                vm = RE_VALUE.search(clean)
                if vm:
                    value = vm.group(1).replace('\\"', '"')
            items.append({"indent": len(indent), "ref": ref, "kind": kind,
                          "name": name, "raw": line, "value": value})
        return items

    @staticmethod
    def norm(text):
        return re.sub(r"[*\s]", "", text or "")

    def find(self, vom, kind, name, nth=0, exact_norm=True):
        hits = [i for i in self.parse(vom)
                if i["kind"] == kind and (self.norm(i["name"]) == self.norm(name) if exact_norm else name in i["name"])]
        if len(hits) <= nth:
            raise StepError(f"页面上找不到 {kind} «{name}»（第 {nth + 1} 个），当前命中 {len(hits)} 个")
        return hits[nth]["ref"]

    def tooltip_options(self, vom, text):
        """在最后一个 tooltip 块里按文本找选项，返回 (refs, 命中总数)。"""
        lines = vom.splitlines()
        start = None
        for idx, line in enumerate(lines):
            if line.strip().startswith("tooltip"):
                start = idx
        if start is None:
            raise StepError("下拉框没有展开（未找到 tooltip 层）")
        indent = len(lines[start]) - len(lines[start].lstrip())
        hits = []
        for line in lines[start + 1:]:
            cur = len(line) - len(line.lstrip())
            if line.strip() and cur <= indent:
                break
            m = RE_REF.match(line)
            if m and m.group(3) == "button":
                name = (m.group(4) or "").replace('\\"', '"')
                if self.norm(name) == self.norm(text):
                    hits.append(m.group(2))
        if not hits:
            raise StepError(f"下拉选项里没有 «{text}»")
        return hits, len(hits)

    # --- 动作 -------------------------------------------------------
    def click(self, ref, ctx):
        self.run(["click", ref, "--session", self.session, "--context", ctx])

    def fill(self, ref, value, ctx):
        self.run(["fill", ref, "--value", value, "--session", self.session, "--context", ctx])

    def js(self, expr, ctx, obs):
        return self.run(["evaluate", expr, "--session", self.session,
                         "--context", ctx, "--observation", obs])

    def screenshot(self, path):
        self.run(["screenshot", "--session", self.session, "--out", path], check=False)


# ---------------------------------------------------------------- 页面业务动作
class TaskForm:
    def __init__(self, br, cfg, log):
        self.br = br
        self.cfg = cfg
        self.log = log

    # --- 浮窗生命周期 ------------------------------------------------
    def close_stray(self):
        for _ in range(3):
            vom = self.br.observe()
            if 'dialog "新增任务"' not in vom:
                return True
            closed = False
            for name in ("el.dialog.close", "关闭窗口"):
                try:
                    ctx, _ = self.br.sync()
                    self.br.click(self.br.find(vom, "button", name), ctx)
                    closed = True
                    break
                except StepError:
                    continue
            if not closed:
                return False
            time.sleep(1.2)
        return 'dialog "新增任务"' not in self.br.observe()

    def form_is_fresh(self, vom):
        """新浮窗的必填输入项必须都是空的，否则说明拿到的是上一个残留表单。"""
        for label in ("任务名称", "任务奖励", "任务按钮跳转链接"):
            try:
                ref = self.br.find(vom, "textbox", label)
            except StepError:
                return False
            val = next(i for i in self.br.parse(vom) if i["ref"] == ref)["value"]
            if self.br.norm(val):
                return False
        return True

    def open_modal(self):
        self.close_stray()
        for attempt in range(3):
            ctx, _ = self.br.sync()
            vom = self.br.observe()
            ref = self.br.find(vom, "button", "新增任务")
            self.br.click(ref, ctx)
            vom = ""
            for _ in range(12):
                time.sleep(0.8)
                vom = self.br.observe()
                if 'dialog "新增任务"' in vom:
                    break
            if 'dialog "新增任务"' in vom and self.form_is_fresh(vom):
                # 浮窗打开后应用可能再执行一次 resetFields，等一拍再确认
                time.sleep(1.2)
                vom = self.br.observe()
                if self.form_is_fresh(vom):
                    return vom
            self.log(f"  浮窗不干净（未打开或带残留值），第 {attempt + 1} 次清理重开", "WARN")
            self.close_stray()
            time.sleep(0.8)
        raise StepError("始终没能打开一个干净的新增任务浮窗")

    def close_modal(self):
        try:
            vom = self.br.observe()
            if 'dialog "新增任务"' not in vom:
                return True
            ctx, _ = self.br.sync()
            self.br.click(self.br.find(vom, "button", "关闭窗口"), ctx)
            time.sleep(1.0)
            return 'dialog "新增任务"' not in self.br.observe()
        except Exception:  # noqa: BLE001
            return False

    # --- 下拉框 ------------------------------------------------------
    def dismiss_popper(self):
        """点标题收起残留下拉框（Esc 会连整个浮窗一起关掉，不能用）。"""
        vom = self.br.observe()
        if not any(line.strip().startswith("tooltip") for line in vom.splitlines()):
            return
        try:
            ctx, _ = self.br.sync()
            self.br.click(self.br.find(vom, "heading", "新增任务"), ctx)
            time.sleep(0.5)
        except StepError:
            pass

    def pick_option(self, label, option, nth=0, retries=3):
        """打开以 label 命名的下拉框并选中 option。等动画收敛再点，失败整轮重来。"""
        last = None
        for attempt in range(retries + 1):
            try:
                self.dismiss_popper()
                ctx, _ = self.br.sync()
                vom = self.br.observe()
                ref = self.br.find(vom, "textbox", label)
                self.br.run(["scroll-to", ref, "--session", self.br.session, "--context", ctx], check=False)
                time.sleep(0.3)
                self.br.click(ref, ctx)
                opened = False
                for _ in range(10):
                    time.sleep(0.4)
                    vom = self.br.observe()
                    if any(l.strip().startswith("tooltip") for l in vom.splitlines()):
                        opened = True
                        break
                if not opened:
                    raise StepError("下拉框没有展开")
                hits, total = self.br.tooltip_options(vom, option)
                if total > 1:
                    self.log(f"  «{label}» 里 «{option}» 有 {total} 个同名选项，取第 {nth + 1} 个", "WARN")
                for k in range(2):
                    try:
                        c, _ = self.br.sync()
                        self.br.click(hits[nth], c)
                        time.sleep(0.5)
                        break
                    except StepError:
                        if k == 1:
                            raise
                        time.sleep(0.6)
                        hits, _ = self.br.tooltip_options(self.br.observe(), option)
                # 选完立刻回读，防止「点到了但没生效」的静默失败
                vom = self.br.observe()
                cur = next(i for i in self.br.parse(vom)
                           if i["ref"] == self.br.find(vom, "textbox", label))["value"]
                if self.br.norm(cur) != self.br.norm(option):
                    raise StepError(f"选中后回读为 «{cur}»，期望 «{option}»")
                return
            except StepError as exc:
                last = exc
                self.log(f"  «{label}» 第 {attempt + 1} 次选择未成功，重试：{str(exc)[:110]}", "WARN")
                time.sleep(0.8)
        raise StepError(f"«{label}» 选择 «{option}» 失败：{last}")

    # --- 输入框 ------------------------------------------------------
    def js_input_value(self, placeholder):
        """直接读可见浮窗的 input.value（VOM 会截断超长文本，长链接必须走这条路）。"""
        expr = ("(()=>{const dl=Array.from(document.querySelectorAll('.el-overlay-dialog'))"
                ".filter(d=>d.getClientRects().length);const d=dl[dl.length-1];if(!d)return 'no-dialog';"
                "const i=Array.from(d.querySelectorAll('input')).find(x=>x.placeholder==="
                + json.dumps(placeholder) + ");return i?i.value:'no-input'})()")
        ctx, obs = self.br.sync()
        out = self.br.js(expr, ctx, obs)
        lines = [l.strip() for l in out.splitlines() if l.strip()]
        return lines[-1] if lines else ""

    def fill_fields(self, row):
        want = [("任务组", row["group"]),
                ("任务名称", row["name"]),
                ("任务按钮跳转链接", row["url"])]
        ctx, _ = self.br.sync()
        vom = self.br.observe()
        plan = [{"action": "fill", "ref": self.br.find(vom, "textbox", label).lstrip("@"), "value": val}
                for label, val in want]
        out = self.br.run(["plan", "--session", self.br.session, "--context", ctx,
                           "--steps", json.dumps(plan, ensure_ascii=False)], check=False)
        if '"ok": true' not in out and '"ok":true' not in out:
            self.log(f"  批量填写返回异常，改为逐项补填：{out.strip()[:160]}", "WARN")

        for label, val in want:
            for _ in range(3):
                if label == "任务按钮跳转链接":
                    if self.js_input_value("请输入任务按钮跳转链接") == val:
                        break
                else:
                    vom = self.br.observe()
                    ref = self.br.find(vom, "textbox", label)
                    cur = next(i for i in self.br.parse(vom) if i["ref"] == ref)["value"]
                    if self.br.norm(cur) == self.br.norm(val):
                        break
                c, _ = self.br.sync()
                self.br.fill(self.br.find(self.br.observe(), "textbox", label), val, c)
                time.sleep(0.4)
            else:
                raise StepError(f"«{label}» 反复填写仍未生效（期望 «{val[:80]}…»）")

    # --- 图片 --------------------------------------------------------
    def upload_image(self, image_url):
        # 关闭过的浮窗会以不可见节点残留，必须锁定「当前可见」的那个
        expr = ("(async()=>{const dlg=Array.from(document.querySelectorAll('.el-overlay-dialog'))"
                ".find(d=>d.getClientRects().length);if(!dlg)return 'no-dialog';"
                "const i=dlg.querySelector('input.el-upload__input');if(!i)return 'no-input';"
                "const r=await fetch(%s);if(!r.ok)return 'fetch-'+r.status;const b=await r.blob();"
                "const dt=new DataTransfer();dt.items.add(new File([b],'task_img.png',{type:'image/png'}));"
                "i.files=dt.files;i.dispatchEvent(new Event('change',{bubbles:true}));return 'ok'})()"
                % json.dumps(image_url))
        for attempt in range(2):
            ctx, obs = self.br.sync()
            out = self.br.js(expr, ctx, obs)
            if "ok" not in out:
                raise StepError(f"图片注入失败：{out.strip()[:200]}")
            for _ in range(20):
                time.sleep(1.2)
                if self.has_image(self.br.observe()):
                    return
            self.log(f"  图片上传未出缩略图，第 {attempt + 1} 次重试", "WARN")
        raise StepError("图片上传后未出现缩略图（可能被后端拒绝）")

    @staticmethod
    def has_image(vom):
        lines = vom.splitlines()
        for idx, line in enumerate(lines):
            if 'group "* 任务图片"' in line:
                base = len(line) - len(line.lstrip())
                for nxt in lines[idx + 1:]:
                    cur = len(nxt) - len(nxt.lstrip())
                    if nxt.strip() and cur <= base:
                        break
                    if nxt.strip().startswith("img"):
                        return True
                return False
        return False

    # --- 校验与提交 ---------------------------------------------------
    def verify(self, row):
        vom = self.br.observe()
        checks = [("任务奖励", row["reward"]), ("任务组", row["group"]),
                  ("任务名称", row["name"]), ("任务类型", row["task_type"]),
                  ("链接类型", row["link_type"]), ("链接所属渠道", row["channel"])]
        for label, expect in checks:
            ref = self.br.find(vom, "textbox", label)
            val = next(i for i in self.br.parse(vom) if i["ref"] == ref)["value"]
            if self.br.norm(val) != self.br.norm(expect):
                raise StepError(f"提交前校验失败：{label} = «{val}»，期望 «{expect}»")
        if self.js_input_value("请输入任务按钮跳转链接") != row["url"]:
            raise StepError("提交前校验失败：任务按钮跳转链接与表格不一致")
        if not self.has_image(vom):
            raise StepError("提交前校验失败：任务图片没有缩略图")
        return vom

    def net_cursor(self):
        out = self.br.run(["network", "--session", self.br.session, "--limit", "5"], check=False)
        nums = [int(x) for x in re.findall(r"#(\d+)", out)]
        return max(nums) if nums else 0

    def toast_text(self):
        out = self.br.run(["read", "--selector", ".el-message", "--json",
                           "--session", self.br.session], check=False)
        texts = re.findall(r'"text":\s*"((?:[^"\\]|\\.)*)"', out)
        return " | ".join(t for t in texts if t.strip())[:200]

    def submit(self, row):
        self.verify(row)
        mark = self.net_cursor()
        ctx, _ = self.br.sync()
        vom = self.br.observe()
        self.br.click(self.br.find(vom, "button", "提交审批"), ctx)
        # 以「本次点击之后是否新增了提交请求」作为唯一成功依据
        committed = False
        for _ in range(12):
            time.sleep(1.0)
            net = self.br.run(["network", "--session", self.br.session,
                               f"--since={mark}", "--limit", "30"], check=False)
            if "commitNewMissionEdit" in net and "POST" in net:
                committed = True
                break
        toast = self.toast_text()
        if not committed:
            raise StepError(f"未检测到提交请求，判定未提交（页面提示：{toast or '无'}）")
        for word in ("失败", "错误", "不能为空", "已存在"):
            if word in toast:
                raise StepError(f"提交返回异常提示：{toast}")
        return toast


# ---------------------------------------------------------------- 主流程
def build_rows(cfg, log):
    cols = {cfg["col_name"], cfg["col_url"], cfg["col_channel"]}
    if cfg.get("col_image"):
        cols.add(cfg["col_image"])
    if cfg.get("col_group"):
        cols.add(cfg["col_group"])

    last_row = cfg.get("end_row")  # 可选：只读到第几行（物理行号，None = 到底）
    raw = read_xlsx_rows(cfg["excel_path"], cfg["sheet_name"], cols, cfg.get("data_start_row", 1))

    excel_dir = os.path.dirname(os.path.abspath(cfg["excel_path"]))
    rows = []
    for n, rno in enumerate(sorted(raw), start=1):
        if last_row and rno > last_row:
            continue
        v = raw[rno]
        img = (v.get(cfg.get("col_image") or "", "") or "").strip()
        if img:
            if not os.path.isabs(img):
                img = os.path.join(excel_dir, img)
        else:
            img = cfg["image_path"]
        rows.append({
            "n": n, "row": rno,
            "name": v.get(cfg["col_name"], ""),
            "url": v.get(cfg["col_url"], ""),
            "channel": v.get(cfg["col_channel"], ""),
            "group": v.get(cfg.get("col_group") or "", "") or cfg["task_group"],
            "reward": cfg["reward_name"],
            "task_type": cfg["task_type"],
            "link_type": cfg["link_type"],
            "image_src": img,
            "image_url": None,
        })
    return rows


def main():
    ap = argparse.ArgumentParser(description="批量创建运营后台任务（行数不限）")
    ap.add_argument("--config", default=os.path.join(ROOT, "config.json"))
    ap.add_argument("--excel"), ap.add_argument("--sheet")
    ap.add_argument("--session"), ap.add_argument("--image")
    ap.add_argument("--only", type=int, help="只跑第 n 条")
    ap.add_argument("--from", dest="frm", type=int)
    ap.add_argument("--to", dest="to", type=int)
    ap.add_argument("--limit", type=int, help="本次最多跑多少条")
    ap.add_argument("--dry-run", action="store_true", help="填完不提交")
    ap.add_argument("--list", action="store_true", help="只打印数据，不碰浏览器")
    ap.add_argument("--force", action="store_true", help="忽略断点记录，强制重跑")
    ap.add_argument("--reset-state", action="store_true", help="清空断点记录后退出")
    ap.add_argument("--continue-on-error", action="store_true",
                    help="单条失败后继续（连续 3 条失败才中止）")
    args = ap.parse_args()

    cfg = load_cfg(args.config)
    for k_cli, k_cfg in (("excel", "excel_path"), ("sheet", "sheet_name"),
                         ("session", "session"), ("image", "image_path")):
        val = getattr(args, k_cli)
        if val:
            cfg[k_cfg] = val

    os.makedirs(cfg["screenshot_dir"], exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    log = Log(os.path.join(cfg["screenshot_dir"], f"run_{stamp}.log"))

    ledger_path = cfg.get("state_path") or os.path.join(ROOT, "state", "submitted.json")
    ledger = Ledger(ledger_path, log)
    if args.reset_state:
        ledger.reset()
        log(f"断点记录已清空：{ledger_path}")
        return 0

    rows = build_rows(cfg, log)
    log(f"表格共 {len(rows)} 行有效数据（{os.path.basename(cfg['excel_path'])} / {cfg['sheet_name']}）")

    # 选段
    if args.only:
        sel = [r for r in rows if r["n"] == args.only]
    else:
        lo = args.frm or 1
        hi = args.to or 10 ** 9
        sel = [r for r in rows if lo <= r["n"] <= hi]

    # 跳过已提交
    done_before = [r for r in sel if ledger.has(r["name"])]
    if done_before and not args.force:
        sel = [r for r in sel if not ledger.has(r["name"])]
        log(f"断点记录里已有 {len(done_before)} 条，自动跳过（--force 可强制重跑）")

    # 跳过空行
    bad = [r["n"] for r in sel if not r["name"] or not r["url"] or not r["channel"]]
    if bad:
        if cfg.get("skip_invalid_rows", True):
            log(f"以下编号缺字段，已跳过：{bad}", "WARN")
            sel = [r for r in sel if r["n"] not in bad]
        else:
            log(f"以下编号缺字段，已中止：{bad}", "ERROR")
            return 2

    if args.limit:
        sel = sel[: args.limit]

    if not sel:
        log("没有需要处理的数据。")
        return 0

    log(f"本次处理 {len(sel)} 条：n={sel[0]['n']}..{sel[-1]['n']}")

    if args.list:
        for r in sel:
            flag = "已提交" if ledger.has(r["name"]) else "待处理"
            log(f"n={r['n']:>3} [{flag}] {r['name']} | {r['channel']} | {os.path.basename(r['image_src'])} | {r['url'][:50]}")
        return 0

    # 预置图片到本地服务目录
    for r in sel:
        r["image_url"] = ensure_served(r["image_src"], cfg["image_dir"], cfg["asset_base_url"])

    if not cfg.get("session"):
        log("没有指定浏览器会话：先 `bsk open <页面地址>` 开一个会话，"
            "再用 --session <id> 传入（或写进 config.json 的 session 字段）", "ERROR")
        return 2

    br = Browser(cfg, log)
    form = TaskForm(br, cfg, log)
    ok, failed, streak = 0, [], 0

    for r in sel:
        log(f"—— n={r['n']} 行{r['row']} «{r['name']}» 开始 ——")
        try:
            form.open_modal()
            form.pick_option("任务奖励", r["reward"], nth=cfg.get("reward_pick_nth", 0))
            form.fill_fields(r)
            form.pick_option("任务类型", r["task_type"])
            form.upload_image(r["image_url"])
            form.pick_option("链接类型", r["link_type"])
            form.pick_option("链接所属渠道", r["channel"])
            form.verify(r)
            if args.dry_run:
                shot = os.path.join(cfg["screenshot_dir"], f"dryrun_n{r['n']}_{stamp}.png")
                br.screenshot(shot)
                log(f"  DRY-RUN 已填完，停在提交前；截图 {shot}", "WARN")
                return 0
            form.submit(r)
            ledger.add(r["name"], {"n": r["n"], "row": r["row"], "channel": r["channel"],
                                   "at": dt.datetime.now().isoformat(timespec="seconds")})
            ok += 1
            streak = 0
            log(f"  ✔ n={r['n']} «{r['name']}» 已提交审批")
        except Exception as exc:  # noqa: BLE001
            shot = os.path.join(cfg["screenshot_dir"], f"fail_n{r['n']}_{stamp}.png")
            br.screenshot(shot)
            log(f"  ✘ n={r['n']} 失败：{exc}", "ERROR")
            log(f"  失败截图：{shot}", "ERROR")
            if form.close_modal():
                log("  已关闭残留浮窗，环境已复原", "WARN")
            else:
                log("  残留浮窗未能确认关闭，下次运行前可能需要人工关掉", "ERROR")
            failed.append(r["n"])
            streak += 1
            if not args.continue_on_error and cfg.get("stop_on_error", True):
                log("按配置中止后续任务（--continue-on-error 可改）。", "ERROR")
                break
            if streak >= 3:
                log("连续 3 条失败，判定环境异常，中止。", "ERROR")
                break

    report = {
        "started": stamp,
        "total_rows": len(rows),
        "processed": len(sel),
        "success": ok,
        "failed": failed,
        "unprocessed": [r["n"] for r in sel[ok + len(failed):]],
    }
    rp = os.path.join(cfg["screenshot_dir"], f"report_{stamp}.json")
    with open(rp, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    log(f"结束：成功 {ok} 条；失败 {failed or '无'}；报告 {rp}")
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
