# -*- coding: utf-8 -*-
"""
统一运营管理平台 —— 浏览任务批量创建脚本（BrowserSkill/bsk 驱动）

流程（每行数据一次）：
  新建任务 -> 任务奖励 -> 任务组 -> 任务名称 -> 任务类型(普通任务-浏览)
  -> 任务图片 -> 任务按钮跳转链接 -> 链接类型(外部跳转) -> 链接所属渠道 -> 提交审批

用法：
  python batch_tasks.py                 # 按 config.json 的 start_index..end_index 全量执行
  python batch_tasks.py --only 1        # 只跑第 1 条（试跑）
  python batch_tasks.py --from 2 --to 50
  python batch_tasks.py --dry-run       # 填完不点提交，停在提交前
  python batch_tasks.py --list          # 只打印将要提交的数据，不碰浏览器
"""
import argparse
import datetime as dt
import json
import os
import re
import subprocess
import sys
import time
import zipfile
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.abspath(__file__))
CFG_PATH = os.path.join(ROOT, "config.json")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

RE_REF = re.compile(r'^(\s*)@(e\d+)\s+([A-Za-z]+)(?:\s+"((?:[^"\\]|\\.)*)")?')
RE_VALUE = re.compile(r'="((?:[^"\\]|\\.)*)"\s*$')


class StepError(RuntimeError):
    pass


# ---------------------------------------------------------------- 配置与日志
def load_cfg():
    with open(CFG_PATH, encoding="utf-8") as f:
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
    """返回 {行号: {列字母: 值}}，不依赖第三方库。"""
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
            raise StepError(f"工作表 {sheet_name} 不存在")
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
                if col not in cols:
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
            if any(vals.get(c) for c in cols):
                rows[rno] = vals
        return rows


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
        start = out.find("{")
        data = json.loads(out[start:])
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
            # 注意：行尾的 ="..." 可能是 placeholder 而不是值，必须先剔除 placeholder
            value = None
            if "[empty]" not in line:
                clean = re.sub(r'placeholder="(?:[^"\\]|\\.)*"', "", line)
                vm = RE_VALUE.search(clean)
                if vm:
                    value = vm.group(1).replace('\\"', '"')
            items.append({
                "indent": len(indent), "ref": ref, "kind": kind, "name": name,
                "raw": line, "value": value,
            })
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
        """在最后一个 tooltip 块里按文本找选项，返回 (ref, 命中总数)。"""
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

    def close_stray(self):
        """关掉所有残留的「新增任务」浮窗，返回是否已无浮窗。"""
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
            # 浮窗渲染有随机延迟，轮询等待出现
            vom = ""
            for _ in range(12):
                time.sleep(0.8)
                vom = self.br.observe()
                if 'dialog "新增任务"' in vom:
                    break
            if 'dialog "新增任务"' in vom and self.form_is_fresh(vom):
                # 浮窗打开后应用可能再执行一次 resetFields，等一拍再确认一次
                time.sleep(1.2)
                vom = self.br.observe()
                if self.form_is_fresh(vom):
                    return vom
            self.log(f"  浮窗不干净（未打开或带残留值），第 {attempt + 1} 次清理重开", "WARN")
            self.close_stray()
            time.sleep(0.8)
        raise StepError("始终没能打开一个干净的新增任务浮窗")

    def dismiss_popper(self):
        """用点击标题的方式收起残留下拉框（Esc 会连整个浮窗一起关掉，不能用）。"""
        vom = self.br.observe()
        if not any(line.strip().startswith("tooltip") for line in vom.splitlines()):
            return
        try:
            ctx, _ = self.br.sync()
            head = self.br.find(vom, "heading", "新增任务")
            self.br.click(head, ctx)
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
                # 等下拉真正展开，避免动画中途取到会移动的元素
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
                for k in range(2):  # 选项点击最多试两次，中间重新观察
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
                # 关键：选完立刻回读，防止「点到了但没生效」的静默失败
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

    def fill_fields(self, row):
        """填任务组/任务名称/跳转链接，并逐项回读，缺哪项补哪项。"""
        want = [("任务组", self.cfg["task_group"]),
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
                vom = self.br.observe()
                ref = self.br.find(vom, "textbox", label)
                cur = next(i for i in self.br.parse(vom) if i["ref"] == ref)["value"]
                if self.br.norm(cur) == self.br.norm(val):
                    break
                c, _ = self.br.sync()
                self.br.fill(ref, val, c)
                time.sleep(0.4)
            else:
                raise StepError(f"«{label}» 反复填写仍未生效（期望 «{val}»）")

    def upload_image(self):
        ctx, obs = self.br.sync()
        expr = ("(async()=>{const r=await fetch(%s);const b=await r.blob();"
                "const dt=new DataTransfer();dt.items.add(new File([b],'task_img.png',{type:'image/png'}));"
                "const i=document.querySelector('.el-dialog input.el-upload__input');if(!i)return 'no-input';"
                "i.files=dt.files;i.dispatchEvent(new Event('change',{bubbles:true}));return 'ok'})()"
                % json.dumps(self.cfg["image_url"]))
        out = self.br.js(expr, ctx, obs)
        if "ok" not in out:
            raise StepError(f"图片注入失败：{out.strip()[:200]}")
        for _ in range(15):
            time.sleep(1.2)
            vom = self.br.observe()
            if self.has_image(vom):
                return
        raise StepError("图片上传后未出现缩略图（可能被后端拒绝）")

    @staticmethod
    def has_image(vom):
        """任务图片分组里是否已出现 img 缩略图。"""
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

    def verify(self, row):
        vom = self.br.observe()
        checks = [("任务奖励", self.cfg["reward_name"]), ("任务组", self.cfg["task_group"]),
                  ("任务名称", row["name"]), ("任务按钮跳转链接", row["url"]),
                  ("任务类型", self.cfg["task_type"]), ("链接类型", self.cfg["link_type"]),
                  ("链接所属渠道", row["channel"])]
        for label, expect in checks:
            ref = self.br.find(vom, "textbox", label)
            val = next(i for i in self.br.parse(vom) if i["ref"] == ref)["value"]
            if self.br.norm(val) != self.br.norm(expect):
                raise StepError(f"提交前校验失败：{label} = «{val}»，期望 «{expect}»")
        if not self.has_image(vom):
            raise StepError("提交前校验失败：任务图片没有缩略图")
        return vom

    def close_modal(self):
        """尽力关闭浮窗，失败不抛异常。"""
        try:
            vom = self.br.observe()
            if 'dialog "新增任务"' not in vom:
                return True
            ctx, _ = self.br.sync()
            ref = self.br.find(vom, "button", "关闭窗口")
            self.br.click(ref, ctx)
            time.sleep(1.0)
            return 'dialog "新增任务"' not in self.br.observe()
        except Exception:  # noqa: BLE001
            return False

    def dialog_hidden(self, timeout="20s"):
        """等待浮窗不可见；用 wait-for-element 而不是脚本，避免观察被页面变化作废。"""
        out = self.br.run(["wait-for-element", "--selector", ".el-dialog", "--state", "hidden",
                           "--timeout", timeout, "--poll-ms", "300", "--json",
                           "--session", self.br.session], check=False)
        return '"satisfied": true' in out or '"satisfied":true' in out, out

    def toast_text(self):
        out = self.br.run(["read", "--selector", ".el-message", "--json",
                           "--session", self.br.session], check=False)
        texts = re.findall(r'"text":\s*"((?:[^"\\]|\\.)*)"', out)
        return " | ".join(t for t in texts if t.strip())[:200]

    def net_cursor(self):
        out = self.br.run(["network", "--session", self.br.session, "--limit", "5"], check=False)
        nums = [int(x) for x in re.findall(r"#(\d+)", out)]
        return max(nums) if nums else 0

    def submit(self, row):
        # 点提交前再完整校验一遍，防止值在最后几秒被应用重置掉
        self.verify(row)
        mark = self.net_cursor()
        ctx, _ = self.br.sync()
        vom = self.br.observe()
        ref = self.br.find(vom, "button", "提交审批")
        self.br.click(ref, ctx)
        # 以「本次点击之后是否真的发出提交请求」作为唯一成功依据
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
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", type=int, help="只跑第 n 条")
    ap.add_argument("--from", dest="frm", type=int)
    ap.add_argument("--to", dest="to", type=int)
    ap.add_argument("--dry-run", action="store_true", help="填完不提交")
    ap.add_argument("--list", action="store_true", help="只打印数据")
    ap.add_argument("--continue-on-error", action="store_true",
                    help="单条失败后继续下一条（连续 3 条失败才中止）")
    args = ap.parse_args()

    cfg = load_cfg()
    os.makedirs(cfg["screenshot_dir"], exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    log = Log(os.path.join(cfg["screenshot_dir"], f"run_{stamp}.log"))

    # 读数据，按 n（1 起）编号
    rows_raw = read_xlsx_rows(cfg["excel_path"], cfg["sheet_name"],
                              {cfg["col_name"], cfg["col_url"], cfg["col_channel"]},
                              cfg["data_start_row"])
    rows = []
    for n, rno in enumerate(sorted(rows_raw), start=1):
        v = rows_raw[rno]
        rows.append({"n": n, "row": rno, "name": v.get(cfg["col_name"], ""),
                     "url": v.get(cfg["col_url"], ""), "channel": v.get(cfg["col_channel"], "")})

    if args.only:
        sel = [r for r in rows if r["n"] == args.only]
    else:
        lo = args.frm or cfg["start_index"]
        hi = args.to or cfg["end_index"]
        sel = [r for r in rows if lo <= r["n"] <= hi]

    log(f"共读取 {len(rows)} 条有效数据，本次处理 {len(sel)} 条（n={sel[0]['n']}..{sel[-1]['n']}）" if sel else "没有可处理的数据")
    missing = [r["n"] for r in sel if not r["name"] or not r["url"] or not r["channel"]]
    if missing:
        log(f"以下编号存在空单元格，已中止：{missing}", "ERROR")
        return 2
    if args.list:
        for r in sel:
            log(f"n={r['n']:>2} 名称={r['name']} | 渠道={r['channel']} | 链接={r['url'][:60]}")
        return 0

    br = Browser(cfg, log)
    form = TaskForm(br, cfg, log)
    ok, failed = 0, []
    streak = 0

    for r in sel:
        log(f"—— n={r['n']} 行{r['row']} «{r['name']}» 开始 ——")
        try:
            form.open_modal()
            form.pick_option("任务奖励", cfg["reward_name"], nth=cfg["reward_pick_nth"])
            form.fill_fields(r)
            form.pick_option("任务类型", cfg["task_type"])
            form.upload_image()
            form.pick_option("链接类型", cfg["link_type"])
            form.pick_option("链接所属渠道", r["channel"])
            form.verify(r)
            if args.dry_run:
                shot = os.path.join(cfg["screenshot_dir"], f"dryrun_n{r['n']}_{stamp}.png")
                br.screenshot(shot)
                log(f"  DRY-RUN 已填完，停在提交前；截图 {shot}", "WARN")
                log("  dry-run 结束，浮窗保持打开，请人工检查后关闭")
                return 0
            form.submit(r)
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
                log("按配置中止后续任务，避免连环错误。", "ERROR")
                break
            if streak >= 3:
                log("连续 3 条失败，判定环境异常，中止。", "ERROR")
                break

    log(f"结束：成功 {ok} 条；失败 {failed or '无'}")
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
