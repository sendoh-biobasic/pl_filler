# -*- coding: utf-8 -*-
"""
PL 数据填充工具（带界面）

做什么：
  打开两个 Excel ——
    ① 导入数据文件：里面有 SAP 导出的明细 sheet（如「导入-8月PL科目数据」）和「COST CENTRE」sheet
    ② 模板文件：里面有要填入的 CA sheet（格式和「导出-PL」一样，带汇总公式）
  把 ① 的数据按规则汇总，只写进 ② 的 CA 里 A 列是 8 位科目代码的「明细行」；
  C 列、小计行、利润行、比率行等已有公式的地方一律不动，用 CA 自己的公式计算。
  「导表-调整」等其他 sheet 完全不碰。结果另存为新文件，不会覆盖原文件。
  可选开关：CA 里 6001、6401、6405 等小计行的各部门列原本取自「导表-调整」的固定数据，
  勾选后改成 =SUM(本组明细行)，部门列的收入/成本/利润就会跟着新数据走。

汇总规则：
  · 第 58~227 行（营业/管理/研发费用三个区块）：先按导入表 Cost Center 分组，再去 COST CENTRE 表
    查成本中心名称判断区块：
        以 -L 结尾 -> 管理费用    Admin/HR -> 管理费用    Finance -> 管理费用
        以 -E 结尾 -> 研发费用
        以 -V 结尾 -> 营业费用    Sales    -> 营业费用
        以 -F 结尾 -> 不处理
  · 其他明细行：A 列科目 = 导入表 G/L Account（所有成本中心相加）。
  · 最后都按导入表 Profit Center 填进对应的列（见 PC_TO_COL）。

运行环境：Python 3 + openpyxl（pip install openpyxl），界面用自带的 tkinter。
"""
import os
import queue
import sys
import threading
import traceback
from collections import defaultdict
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

# ======================= 配置区 =======================
DEFAULT_TARGET_SHEET = 'CA'          # 模板里要填的 sheet
FIRST_ROW = 7                        # 第一个科目行

# 导入表 Profit Center  ->  CA 的列
PC_TO_COL = {
    '3201001101': 'D',    # 合成
    '3201002101': 'E',    # 基因
    '3202001101': 'I',    # 测序
    '3203001101': 'N',    # 试剂
    '3203002101': 'O',    # 试剂盒
    '3203003101': 'P',    # 耗材
    '3208001101': 'Q',    # 蛋白
    '3209002101': 'R',    # 蛋白服务
    '3208003101': 'S',    # 抗体
    '3208004101': 'T',    # 抗体服务
    '3209005101': 'U',    # 多肽
    '3209010301': 'V',    # 行政
    '3209020301': 'AA',   # 财务
}

# 科目代码以这些数字开头的行要「取反」（SAP 金额 × -1）。
# 已用 导出-PL 现有数字验证：6001、6701、6702、66030002；
# 6051 6101 6111 6115 6117 6301 目前导入数据里没有金额，按「收入/收益为正」惯例放进来，首次出现金额时请核对符号。
SIGN_FLIP_PREFIXES = (
    '6001',
    '6051', '6101', '6111', '6115', '6117', '6301',
    '6701', '6702',
    '66030002',          # 财务费用-利息收入（CA 里的财务费用公式是 =利息收支 - 利息收入，所以显示为正）
)

# 三个费用区块 {区块名: (首行, 末行, 汇总行)}；汇总行 B 列必须是 '<区块名>费用'
EXPENSE_BLOCKS = {
    '营业': (58, 113, 114),
    '管理': (115, 170, 171),
    '研发': (172, 227, 228),
}
SECTION_BY_SUFFIX = {'L': '管理', 'E': '研发', 'V': '营业', 'F': None}   # None = 不处理
SECTION_BY_NAME = {'Admin/HR': '管理', 'Sales': '营业', 'Finance': '管理'}
# =====================================================


class FillError(Exception):
    """给用户看的错误信息。"""


def norm_code(v):
    """科目/利润中心/成本中心代码统一成字符串（表里有的是文本、有的是数字）。"""
    if v is None:
        return ''
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return str(int(v))
    return str(v).strip()


def find_cost_centre_sheet(names):
    for n in names:
        key = n.upper().replace(' ', '')
        if key.startswith('COSTCENTRE') or key.startswith('COSTCENTER'):
            return n
    return None


# ---------------------- 读取导入数据 ----------------------
def read_import(path, sheet_name, log):
    """返回 ({(科目, 成本中心, 利润中心): 金额}, {成本中心代码: 名称})。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise FillError(f'导入数据文件里没有「{sheet_name}」这个 sheet。')
        rows = wb[sheet_name].iter_rows(values_only=True)
        header = [str(h).strip() if h is not None else '' for h in next(rows)]
        need = ('G/L Account', 'Amount in Local Currency', 'Cost Center', 'Profit Center')
        missing = [h for h in need if h not in header]
        if missing:
            raise FillError(f'「{sheet_name}」第 1 行找不到这些列标题：{", ".join(missing)}')
        i_gl, i_amt, i_cc, i_pc = (header.index(h) for h in need)

        sums = defaultdict(float)
        used = skipped = 0
        for r in rows:
            gl, amt = norm_code(r[i_gl]), r[i_amt]
            if not gl.isdigit() or isinstance(amt, bool) or not isinstance(amt, (int, float)):
                skipped += 1            # 表头重复行、'Account xxx' 小计行等
                continue
            sums[(gl, norm_code(r[i_cc]), norm_code(r[i_pc]))] += amt
            used += 1
        log(f'导入数据：使用 {used} 行，跳过非数据行 {skipped} 行')

        cc_sheet = find_cost_centre_sheet(wb.sheetnames)
        if cc_sheet is None:
            raise FillError('导入数据文件里找不到「COST CENTRE」sheet（用来判断成本中心属于哪个区块）。')
        cost_centres = {}
        for a, b, *_ in wb[cc_sheet].iter_rows(min_row=2, max_col=2, values_only=True):
            if a is not None and b is not None:
                cost_centres[norm_code(a)] = str(b).strip()
        log(f'成本中心对照表：{len(cost_centres)} 个成本中心')
        return sums, cost_centres
    finally:
        wb.close()


# ---------------------- 汇总 ----------------------
def section_of(name):
    """成本中心名称 -> (区块, 是否认识)。区块为 None 表示不处理。"""
    if name in SECTION_BY_NAME:
        return SECTION_BY_NAME[name], True
    suffix = name.rsplit('-', 1)[-1] if '-' in name else None
    if suffix in SECTION_BY_SUFFIX:
        return SECTION_BY_SUFFIX[suffix], True
    return None, False


def aggregate(sums, cost_centres, problems):
    """返回 (block_agg{(区块,科目,利润中心):金额}, plain_agg{(科目,利润中心):金额})。"""
    block_agg, plain_agg = defaultdict(float), defaultdict(float)
    for (gl, cc, pc), v in sums.items():
        if abs(v) <= 0.005:
            continue
        if not gl.startswith('6002'):
            plain_agg[(gl, pc)] += v
            continue
        # 6002xxxx 费用科目：先按成本中心判断区块
        if not cc:
            problems.append(f'科目 {gl} 利润中心 {pc or "(空)"} 金额 {v:,.2f}：没有成本中心，无法判断区块')
            continue
        name = cost_centres.get(cc)
        if name is None:
            problems.append(f'科目 {gl} 成本中心 {cc} 金额 {v:,.2f}：COST CENTRE 表里没有这个成本中心')
            continue
        sec, known = section_of(name)
        if not known:
            problems.append(f'科目 {gl} 成本中心 {cc}（{name}）金额 {v:,.2f}：名称不符合任何规则（-L/-E/-V/-F/Admin/HR/Sales/Finance）')
            continue
        if sec is None:                  # -F：不处理
            continue
        block_agg[(sec, gl, pc)] += v
    return block_agg, plain_agg


# ---------------------- 写入模板 ----------------------
def check_layout(ws):
    """确认模板的结构和预期一致，否则宁可停下也不乱写。"""
    for sec, (first, last, sub) in EXPENSE_BLOCKS.items():
        label = str(ws.cell(sub, 2).value or '').strip()
        if label != f'{sec}费用':
            raise FillError(f'「{ws.title}」第 {sub} 行 B 列应为「{sec}费用」，实际是「{label}」。'
                            f'表格结构和预期不一致，已停止。')
    label = str(ws.cell(229, 2).value or '').strip()
    if label != '营业、管理与研发费用合计':
        raise FillError(f'「{ws.title}」第 229 行 B 列应为「营业、管理与研发费用合计」，实际是「{label}」，已停止。')


def put_value(ws, coord, value):
    """写入数值；原来是公式/手工数字都会被替换；空格子借用同行 D 或 C 列的格式。返回原来的内容类型。"""
    cell = ws[coord]
    old = cell.value
    kind = 'empty' if old is None else ('formula' if isinstance(old, str) and old.startswith('=') else 'value')
    if not cell.has_style:
        for src_col in (4, 3):
            src = ws.cell(cell.row, src_col)
            if src.has_style:
                cell._style = copy(src._style)
                break
    cell.value = value
    return kind


def find_subtotal_rows(ws):
    """找出 CA 里「4 位科目代码、且各部门列取自 导表-调整」的小计行，
    返回 [(行号, 代码, 明细首行, 明细末行)]。明细行 = 紧挨在它上面、代码以该 4 位数字开头的连续明细行
    （费用区块的汇总行直接用区块范围）。"""
    block_sub = {sub: (a, b) for a, b, sub in EXPENSE_BLOCKS.values()}
    result = []
    for r in range(FIRST_ROW, ws.max_row + 1):
        code = norm_code(ws.cell(r, 1).value)
        if not (code.isdigit() and len(code) == 4):
            continue
        refs = any(isinstance(ws[f'{c}{r}'].value, str) and '导表-调整' in ws[f'{c}{r}'].value
                   for c in PC_TO_COL.values())
        if not refs:
            continue
        if r in block_sub:
            first, last = block_sub[r]
        else:
            last = r - 1
            first = last
            while first - 1 >= FIRST_ROW and norm_code(ws.cell(first - 1, 1).value).startswith(code) \
                    and len(norm_code(ws.cell(first - 1, 1).value)) == 8:
                first -= 1
            if not norm_code(ws.cell(last, 1).value).startswith(code):
                continue
        result.append((r, code, first, last))
    return result


def fill_sheet(ws, block_agg, plain_agg, log, problems, rebuild_subtotals=False):
    check_layout(ws)
    block_of_row = {r: sec for sec, (a, b, _) in EXPENSE_BLOCKS.items() for r in range(a, b + 1)}

    leaf_rows = []
    for r in range(FIRST_ROW, ws.max_row + 1):
        code = norm_code(ws.cell(r, 1).value)
        if code.isdigit() and len(code) == 8:
            leaf_rows.append((r, code))

    stats = defaultdict(int)
    plain_codes, block_codes = set(), set()
    for r, code in leaf_rows:
        sec = block_of_row.get(r)
        (block_codes.add((sec, code)) if sec else plain_codes.add(code))
        sign = -1 if code.startswith(SIGN_FLIP_PREFIXES) else 1
        for pc, col in PC_TO_COL.items():
            raw = block_agg.get((sec, code, pc), 0.0) if sec else plain_agg.get((code, pc), 0.0)
            kind = put_value(ws, f'{col}{r}', round(raw * sign, 2) + 0.0)
            stats[kind] += 1
        if ws.cell(r, 3).value is None:            # 个别明细行 C 列是空的：补上和其他行一样的求和
            put_value(ws, f'C{r}', f'=SUM(D{r}:AA{r})')
            stats['C补公式'] += 1

    # 检查：有金额但没地方放的数据
    for (sec, gl, pc), v in sorted(block_agg.items()):
        if (sec, gl) not in block_codes:
            problems.append(f'科目 {gl}（{sec}费用）金额 {v:,.2f}：该区块里没有这个科目的行，未填入')
        elif pc not in PC_TO_COL:
            problems.append(f'科目 {gl}（{sec}费用）利润中心 {pc or "(空)"} 金额 {v:,.2f}：利润中心没有对应的列，未填入')
    for (gl, pc), v in sorted(plain_agg.items()):
        if gl not in plain_codes:
            problems.append(f'科目 {gl} 金额 {v:,.2f}：模板里没有这个科目的明细行，未填入')
        elif pc not in PC_TO_COL:
            problems.append(f'科目 {gl} 利润中心 {pc or "(空)"} 金额 {v:,.2f}：利润中心没有对应的列，未填入')

    log(f'已处理 {len(leaf_rows)} 个明细行 × {len(PC_TO_COL)} 个利润中心列 = {len(leaf_rows) * len(PC_TO_COL)} 个单元格')
    log(f'  其中替换了原来的公式 {stats["formula"]} 个、手工输入的数字 {stats["value"]} 个，原本为空 {stats["empty"]} 个')
    if stats['C补公式']:
        log(f'  另有 {stats["C补公式"]} 个明细行的 C 列是空的，已补上 =SUM(D:AA)')

    subs = find_subtotal_rows(ws)
    if rebuild_subtotals:
        for r, code, first, last in subs:
            for col in PC_TO_COL.values():
                put_value(ws, f'{col}{r}', f'=SUM({col}{first}:{col}{last})')
        log(f'小计行改成 SUM 公式：' + '、'.join(f'{code}(第{r}行)' for r, code, _, _ in subs))
    elif subs:
        log('[提示] 以下小计行的各部门列仍取自「导表-调整」的固定数据，没有跟着新数据变（C 总计列是对的）：')
        log('       ' + '、'.join(f'{code}(第{r}行)' for r, code, _, _ in subs))
        log('       这会让这些行以及下面的利润/比率行的部门列和新数据对不上；勾选「小计行也改成 SUM 公式」可以解决。')
    return stats


def run_fill(import_path, import_sheet, template_path, target_sheet, out_path, log, rebuild_subtotals=False):
    import_path, template_path, out_path = Path(import_path), Path(template_path), Path(out_path)
    for label, p in (('导入数据文件', import_path), ('模板文件', template_path)):
        if not p.exists():
            raise FillError(f'找不到{label}：{p}')
    if out_path.resolve() == template_path.resolve() or out_path.resolve() == import_path.resolve():
        raise FillError('保存位置不能和原文件相同，请换一个文件名（程序不会覆盖原文件）。')

    sums, cost_centres = read_import(import_path, import_sheet, log)
    problems = []
    block_agg, plain_agg = aggregate(sums, cost_centres, problems)

    log('正在打开模板（文件较大时需要几秒）…')
    wb = load_workbook(template_path, keep_vba=template_path.suffix.lower() == '.xlsm')
    if target_sheet not in wb.sheetnames:
        raise FillError(f'模板文件里没有「{target_sheet}」这个 sheet。')
    fill_sheet(wb[target_sheet], block_agg, plain_agg, log, problems, rebuild_subtotals)

    wb.calculation.fullCalcOnLoad = True          # 用 Excel 打开时自动重新计算所有公式
    wb.save(out_path)
    log(f'\n已保存：{out_path}')
    log('用 Excel 打开后公式会自动重新计算。')
    if problems:
        log('\n[注意] 以下数据没有被填入，请检查：')
        for p in problems:
            log('  ' + p)
    else:
        log('\n导入的所有金额都已找到位置，没有遗漏。')
    return out_path, problems


# ---------------------- 界面 ----------------------
def build_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText

    try:                                           # Windows 高分屏下文字更清晰
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title('PL 数据填充工具')
    root.geometry('820x560')
    root.minsize(720, 480)
    root.columnconfigure(1, weight=1)

    v_imp, v_imp_sheet = tk.StringVar(), tk.StringVar()
    v_tpl, v_tpl_sheet = tk.StringVar(), tk.StringVar()
    v_out = tk.StringVar()
    v_sub = tk.BooleanVar(value=False)
    q = queue.Queue()
    result = {'out': None}

    def sheet_names(path):
        wb = load_workbook(path, read_only=True)
        try:
            return list(wb.sheetnames)
        finally:
            wb.close()

    def pick_import():
        p = filedialog.askopenfilename(title='选择导入数据文件', filetypes=[('Excel', '*.xlsx *.xlsm')])
        if not p:
            return
        v_imp.set(p)
        try:
            names = sheet_names(p)
        except Exception as e:
            messagebox.showerror('打开失败', str(e))
            return
        cb_imp['values'] = names
        v_imp_sheet.set(next((n for n in names if '导入' in n), names[0]))

    def pick_template():
        p = filedialog.askopenfilename(title='选择模板文件（含 CA 表）', filetypes=[('Excel', '*.xlsx *.xlsm')])
        if not p:
            return
        v_tpl.set(p)
        try:
            names = sheet_names(p)
        except Exception as e:
            messagebox.showerror('打开失败', str(e))
            return
        cb_tpl['values'] = names
        v_tpl_sheet.set(DEFAULT_TARGET_SHEET if DEFAULT_TARGET_SHEET in names else names[0])
        tp = Path(p)
        v_out.set(str(tp.with_name(tp.stem + '_已填充' + tp.suffix)))

    def pick_out():
        init = Path(v_out.get()) if v_out.get() else None
        p = filedialog.asksaveasfilename(title='结果另存为', defaultextension='.xlsx',
                                         initialfile=init.name if init else '',
                                         initialdir=str(init.parent) if init else None,
                                         filetypes=[('Excel', '*.xlsx *.xlsm')])
        if p:
            v_out.set(p)

    pad = {'padx': 8, 'pady': 5}
    ttk.Label(root, text='① 导入数据文件').grid(row=0, column=0, sticky='w', **pad)
    ttk.Entry(root, textvariable=v_imp).grid(row=0, column=1, sticky='ew', **pad)
    ttk.Button(root, text='浏览…', command=pick_import).grid(row=0, column=2, **pad)
    ttk.Label(root, text='   数据所在 sheet').grid(row=1, column=0, sticky='w', **pad)
    cb_imp = ttk.Combobox(root, textvariable=v_imp_sheet, state='readonly')
    cb_imp.grid(row=1, column=1, sticky='ew', **pad)

    ttk.Label(root, text='② 模板文件').grid(row=2, column=0, sticky='w', **pad)
    ttk.Entry(root, textvariable=v_tpl).grid(row=2, column=1, sticky='ew', **pad)
    ttk.Button(root, text='浏览…', command=pick_template).grid(row=2, column=2, **pad)
    ttk.Label(root, text='   要填入的 sheet').grid(row=3, column=0, sticky='w', **pad)
    cb_tpl = ttk.Combobox(root, textvariable=v_tpl_sheet, state='readonly')
    cb_tpl.grid(row=3, column=1, sticky='ew', **pad)

    ttk.Label(root, text='③ 结果保存为').grid(row=4, column=0, sticky='w', **pad)
    ttk.Entry(root, textvariable=v_out).grid(row=4, column=1, sticky='ew', **pad)
    ttk.Button(root, text='另存为…', command=pick_out).grid(row=4, column=2, **pad)

    ttk.Checkbutton(root, variable=v_sub,
                    text='小计行（6001、6401、6405 等）也改成 SUM 公式（不勾选则这些行的部门列继续取「导表-调整」的固定数据）'
                    ).grid(row=5, column=0, columnspan=3, sticky='w', **pad)
    btn = ttk.Button(root, text='开始填充')
    btn.grid(row=6, column=0, columnspan=3, pady=8)
    root.rowconfigure(7, weight=1)
    txt = ScrolledText(root, height=16, state='disabled', font=('Consolas', 10))
    txt.grid(row=7, column=0, columnspan=3, sticky='nsew', padx=8, pady=(0, 8))
    btn_open = ttk.Button(root, text='打开结果所在文件夹', state='disabled')
    btn_open.grid(row=8, column=0, columnspan=3, pady=(0, 8))

    def write(msg):
        txt.configure(state='normal')
        txt.insert('end', msg + '\n')
        txt.see('end')
        txt.configure(state='disabled')

    def poll():
        try:
            while True:
                kind, payload = q.get_nowait()
                if kind == 'log':
                    write(payload)
                elif kind == 'done':
                    btn.configure(state='normal')
                    result['out'] = payload
                    btn_open.configure(state='normal')
                    messagebox.showinfo('完成', f'已保存：\n{payload}')
                elif kind == 'error':
                    btn.configure(state='normal')
                    write('\n[出错] ' + payload)
                    messagebox.showerror('出错', payload)
        except queue.Empty:
            pass
        root.after(100, poll)

    def worker(args):
        try:
            out, _ = run_fill(*args[:5], log=lambda m: q.put(('log', m)), rebuild_subtotals=args[5])
            q.put(('done', str(out)))
        except FillError as e:
            q.put(('error', str(e)))
        except Exception:
            q.put(('error', '程序出错了，请把下面的信息发给开发者：\n' + traceback.format_exc()))

    def start():
        if not (v_imp.get() and v_tpl.get() and v_out.get()):
            messagebox.showwarning('还没选完', '请先选好导入数据文件、模板文件和保存位置。')
            return
        txt.configure(state='normal')
        txt.delete('1.0', 'end')
        txt.configure(state='disabled')
        btn.configure(state='disabled')
        btn_open.configure(state='disabled')
        threading.Thread(target=worker, daemon=True,
                         args=((v_imp.get(), v_imp_sheet.get(), v_tpl.get(), v_tpl_sheet.get(), v_out.get(), v_sub.get()),)).start()

    def open_folder():
        if result['out']:
            folder = str(Path(result['out']).parent)
            try:
                os.startfile(folder)                  # Windows
            except AttributeError:
                import subprocess
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', folder])

    btn.configure(command=start)
    btn_open.configure(command=open_folder)
    root.after(100, poll)
    return root


if __name__ == '__main__':
    build_gui().mainloop()