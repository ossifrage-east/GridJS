/*  # 工作表命名空间（Python 桥）TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * 工作表命名空间：Python 脚本与表格数据之间的双向桥
 *
 * ## 职责划分
 *
 * 这里只做「数据搬运」，不含任何 Pyodide 相关代码（Pyodide 的加载与执行在
 * `src/utils/python/` 下），因此本模块可被主线程直接使用：
 *
 * 1. {@link PY_NAMESPACE_SOURCE} —— 注入到 Pyodide 的 Python 模块 `silkspaces` 源码。
 *    它由主线程随每次运行请求下发给 Worker（而不是让 Worker `import` 本模块），
 *    否则 Worker bundle 会被本模块依赖的 `DataCollection` / `Cell` 整条数据层拖进来。
 * 2. {@link buildSnapshot} —— 运行前把当前工作表的非空格子导出为 JSON 快照。
 * 3. {@link applyWrites} —— 运行后把脚本的写回改动套进撤销栈落到表格上。
 *
 * ## 为什么两边都传 JSON 字符串
 *
 * Pyodide 的 `globals.set` 会把 JS 对象包成 `JsProxy`，Python 侧需要显式
 * `.destroy()` 才不泄漏；JSON 字符串是纯值，两端各解析/序列化一次即可，
 * 完全规避代理的生命周期问题。写回同理（Worker 回传 JSON 字符串）。
 *
 * @module sheetNameSpace
 */

import { DataCollection } from '../../core/dataArchitecture/dataCollection';
import { Cell } from '../../core/dataArchitecture/cell';

/**
 * `silkspaces` Python 模块源码
 *
 * ⚠️ 本字符串是 TS 模板字面量，因此**源码里刻意不出现反斜杠、反引号与 `${`**：
 * JS 模板字面量会把 `\$` 折成 `$`、`\s` 折成 `s`，稍不留神就会静默改坏正则；
 * 反引号更狠 —— 它会**提前终止**本字面量，其后整段被当成 JS 代码解析，
 * 报错却指向字符串起点（形如 TS2349「This expression is not callable」）。
 * 故正则只用字符类、范围解析改成 `split(':')`、格式化只用 `%`。
 *
 * 模块内以 `_` 开头的名字（`_load` / `_dump` / `_parse` …）不会被用户脚本的
 * `from silkspaces import *` 带进命名空间，属于宿主与模块之间的内部接口。
 */
export const PY_NAMESPACE_SOURCE = `
"""SilkSpaces 工作表命名空间。

「脚本」页签里的脚本可以通过这里的函数读写**当前工作表**：

    print(sheet_name, row_count, col_count)
    print(get_cell("A1"))
    set_cell("A1", "Hello")
    set_range("B2:C3", [["b2", "c2"], ["b3", "c3"]])

也可以把**单元格名直接当变量名**用（列字母须大写），模块级与函数体内都成立：

    A1 = 1                    # 等价于 set_cell("A1", 1)
    C10 = "2026-10-20"        # 等价于 set_cell("C10", "2026-10-20")
    x = A1                    # 等价于 x = get_cell("A1")

    def main():
        B2 = "函数体里也可以"  # 同样写入 B2

做法见 _CellRewrite：脚本编译前会把这类标识符改写成 __cells__["A1"]，
于是「函数体内的赋值是局部变量、根本碰不到工作表」的 Python 作用域问题不再存在。

脚本运行结束后，写入的改动由宿主一次性回写到表格，并且可以 Ctrl+Z 撤销。
"""

import ast as _ast
import json as _json
import re as _re

# 当前工作表快照 { "A1": "文本" }，由宿主注入
_SNAPSHOT = {}
# 本次运行产生的写入 { "A1": "新值" }（空串表示清空）
_WRITES = {}

# 当前工作表名（即脚本的存储键）
sheet_name = ""
# 当前表格行数
row_count = 0
# 当前表格列数
col_count = 0

_CELL_NAME = _re.compile(r"^([A-Za-z]{1,3})([0-9]{1,7})$")


def _col_index(letters):
    """列字母换算为列号（A=1，Z=26，AA=27）。"""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _col_name(index):
    """列号换算为列字母。"""
    name = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def _parse(name):
    """解析单元格名，返回 (列号, 行号)，均为 1 起始；允许小写与 $ 前缀。"""
    cleaned = name.strip().replace("$", "")
    m = _CELL_NAME.match(cleaned)
    if not m:
        raise ValueError("不是合法的单元格名：%s" % (name,))
    return _col_index(m.group(1)), int(m.group(2))


def _name(col, row):
    """由列号、行号生成规范单元格名（如 A1）。"""
    return "%s%d" % (_col_name(col), row)


def _parse_range(a1):
    """解析 "A1:C3" 或单格 "A1"，返回按输入顺序的 (c1, r1, c2, r2)。"""
    parts = [p for p in a1.strip().split(":") if p.strip()]
    if len(parts) == 1:
        col, row = _parse(parts[0])
        return col, row, col, row
    c1, r1 = _parse(parts[0])
    c2, r2 = _parse(parts[-1])
    return c1, r1, c2, r2


def _key_order(key):
    """排序键：先行后列，使 iter_cells 的遍历顺序与表格视觉顺序一致。"""
    col, row = _parse(key)
    return (row, col)


def _load(snapshot_json, name, rows, cols):
    """宿主在运行脚本前调用：注入工作表快照（内部接口）。"""
    global _SNAPSHOT, _WRITES, sheet_name, row_count, col_count
    _SNAPSHOT = _json.loads(snapshot_json)
    _WRITES = {}
    __cells__.reset()
    sheet_name = name
    row_count = rows
    col_count = cols


def _dump():
    """宿主在运行脚本后调用：取出本次写入（内部接口）。"""
    return _json.dumps(_WRITES)


# 变量名形态的单元格标识：**列字母大写** + 行号（A1、C10、AA100）。
# 刻意只认大写：x1 / b2 / a1 这类小写名在 Python 里太常见，
# 若也当成单元格名，普通脚本会被静默改成「往表里写值」，代价远大于收益。
# 函数式 API（get_cell / set_cell）仍照旧大小写不限，见 _CELL_NAME。
_VAR_NAME = _re.compile(r"^[A-Z]{1,3}[0-9]{1,7}$")


class _Cells:
    """单元格命名空间：脚本里 A1 = 1 这类写法最终都落到这里（内部接口）。

    ## 为什么不直接拦截用户脚本的 globals

    A1 = 1 在**模块级**编译为 STORE_NAME、在**函数体内**编译为 STORE_FAST。
    后者直接写进帧对象的局部数组，任何字典子类的 __setitem__ 都拦不到；函数内
    再加 global 声明也无效（STORE_GLOBAL 走的是 C 层 PyDict_SetItem，同样绕过
    Python 层钩子）。实测确认过这一点，故改为在编译前用 AST 改写（见 _CellRewrite）。

    ## 读与写的分工

    - 写：__setitem__ 记入 _WRITES（表层语义，str 化；None 等价清空），
      同时把**原始 Python 值**留一份在 _vals，使 A1 + 1 这类运算仍按数值进行；
    - 读：优先取本次运行赋过的原始值，其次取 _WRITES，最后取快照文本；
    - 删：del A1 等价 clear_cell（_WRITES 置空串，原始值一并丢弃）。
    """

    def __init__(self):
        self._vals = {}

    def __getitem__(self, key):
        if key in self._vals:
            return self._vals[key]
        if key in _WRITES:
            return _WRITES[key]
        return _SNAPSHOT.get(key, "")

    def __setitem__(self, key, value):
        self._vals[key] = value
        _WRITES[key] = "" if value is None else str(value)

    def __delitem__(self, key):
        self._vals.pop(key, None)
        _WRITES[key] = ""

    def reset(self):
        """宿主在每次运行前清空（与 _load 一同调用）。"""
        self._vals.clear()


__cells__ = _Cells()


class _CellRewrite(_ast.NodeTransformer):
    """把「单元格名形态」的标识符改写成 __cells__["A1"]（内部接口）。

    改写成下标后，A1 = 1 / x = A1 / A1 += 1 / A1, B1 = 1, 2 /
    for A1 in ... / del A1 全都自动成立 —— 下标本身就是合法的赋值目标，
    不必为每种语法各写一条规则，也不必区分当前是模块级还是函数体。

    唯一要排除的是**函数 / Lambda 的形参同名**（如 def f(A1): return A1 * 2）：
    那是普通局部变量，不该被接到单元格上。
    """

    def __init__(self):
        self._params = []

    def _arg_names(self, args):
        names = [a.arg for a in args.posonlyargs + args.args + args.kwonlyargs]
        if args.vararg:
            names.append(args.vararg.arg)
        if args.kwarg:
            names.append(args.kwarg.arg)
        return names

    def _enter(self, args):
        self._params.append(set(self._arg_names(args)))

    def _leave(self):
        self._params.pop()

    def visit_FunctionDef(self, node):
        self._enter(node.args)
        self.generic_visit(node)
        self._leave()
        return node

    def visit_AsyncFunctionDef(self, node):
        self._enter(node.args)
        self.generic_visit(node)
        self._leave()
        return node

    def visit_Lambda(self, node):
        self._enter(node.args)
        self.generic_visit(node)
        self._leave()
        return node

    def visit_Name(self, node):
        if not _VAR_NAME.match(node.id):
            return node
        for bound in self._params:
            if node.id in bound:
                return node
        cell = _ast.Subscript(
            value=_ast.Name(id="__cells__", ctx=_ast.Load()),
            slice=_ast.Constant(value=node.id),
            ctx=node.ctx,
        )
        return _ast.copy_location(cell, node)


def _compile(source, filename):
    """编译用户脚本：先做单元格名改写，再交给内置 compile（内部接口）。

    copy_location 把新节点的行列信息对齐到原节点，故 traceback 里的行号
    与用户脚本一致；fix_missing_locations 只兜底补上极少数缺位置的新节点。
    """
    tree = _CellRewrite().visit(_ast.parse(source, filename, "exec"))
    _ast.fix_missing_locations(tree)
    return compile(tree, filename, "exec")


def get_cell(name):
    """取单元格文本；空格子返回空串，本次已写入的值优先于快照。"""
    col, row = _parse(name)
    key = _name(col, row)
    if key in _WRITES:
        return _WRITES[key]
    return _SNAPSHOT.get(key, "")


def set_cell(name, value):
    """设置单元格文本；value 为 None 时等价清空；其余按 str(value) 写入。"""
    col, row = _parse(name)
    _WRITES[_name(col, row)] = "" if value is None else str(value)


def clear_cell(name):
    """清空单元格。"""
    col, row = _parse(name)
    _WRITES[_name(col, row)] = ""


def get_range(a1):
    """取 "A1:C3" 区域的文本，返回二维列表（行 × 列，越界补空串）。"""
    c1, r1, c2, r2 = _parse_range(a1)
    out = []
    for row in range(min(r1, r2), max(r1, r2) + 1):
        out.append([get_cell(_name(col, row)) for col in range(min(c1, c2), max(c1, c2) + 1)])
    return out


def set_range(a1, values):
    """写入区域：以 a1 为左上角，把 values 逐格写入。

    values 为二维序列（一维序列视为一行）；不做裁剪，多余的行/列会继续
    向右下方向写入。单格 a1（如 "A1"）时按一行处理。
    """
    rows = list(values)
    if rows and isinstance(rows[0], (list, tuple)):
        rows = [list(row) for row in rows]
    else:
        rows = [rows]
    c1, r1, _c2, _r2 = _parse_range(a1)
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            set_cell(_name(c1 + j, r1 + i), value)


def iter_cells():
    """迭代当前工作表所有非空格子，产出 (名称, 文本)；含本次写入的改动。"""
    merged = dict(_SNAPSHOT)
    merged.update(_WRITES)
    for key in sorted(merged, key=_key_order):
        text = merged[key]
        if text:
            yield key, text
`;

/**
 * 当前工作表的可运行快照
 *
 * {@link json} 是发给 Python 的载荷（`{"A1": "文本", ...}`）；{@link cells} 与
 * {@link json} 内容等价，保留数组形式只是为了便于调试与统计。
 */
export interface SheetSnapshot {
    /** 当前工作表名 */
    sheetName: string;
    /** 表格行数 */
    rowCount: number;
    /** 表格列数 */
    colCount: number;
    /** 非空格子（`[单元格名, 文本]`） */
    cells: [string, string][];
    /** 发给 Python 的 JSON 载荷 */
    json: string;
}

/**
 * 导出当前工作表的快照（运行时送给 Python 的输入）
 *
 * 只带**非空**格子：空串在 Python 侧与「不存在」等价（`get_cell` 返回 `""`），
 * 而百万行大表里空格子往往占绝大多数，带上它们只是徒增序列化成本。
 * @param data - 主数据集合
 * @returns {SheetSnapshot} 快照
 */
export function buildSnapshot(data: DataCollection): SheetSnapshot {
    const cells: [string, string][] = [];
    const map: Record<string, string> = {};
    for (const cell of data.values) {
        const text = cell.getText();
        if (text === '') continue;
        const name = cell.cell;
        cells.push([name, text]);
        map[name] = text;
    }
    return {
        sheetName: data.activedSheetName ?? '',
        rowCount: data.rowHeaders.length,
        colCount: data.colHeaders.length,
        cells,
        json: JSON.stringify(map),
    };
}

/**
 * 把脚本的写回改动落到当前工作表上（纳入撤销栈）
 *
 * ## 为什么走 runWithCellsUndo 而不是 runWithCellPatchUndo
 *
 * 宏脚本的写入是**稀疏**的（典型是个位数的格子），两个「局部补丁」撤销 API 中：
 * - `runWithCellPatchUndo` 按矩形逐格捕获快照，代价与矩形面积成正比 —— 若脚本
 *   同时写了 A1 与 Z9999，就要为 26 万格各克隆一份（几百毫秒 + 成百 MB）；
 * - `runWithCellsUndo` 按给定数组逐个捕获，代价只与写入格数成正比
 *   （查找替换用的就是它）。
 *
 * 因此统一走稀疏版。唯一的语义边界：本次**新建**的空格子其「操作前」快照是空内容
 * 而非 `null`，撤销时该空格子会留在 `values` 里，由下一次 `commitValues()` 的
 * `_pruneEmptyCells()` 清掉 —— 属该 API 的既有语义，代价可忽略。
 *
 * `runWithCellsUndo` 内部已含「清扫空格 → 标记内容变更 → 入撤销栈 → 广播
 * VALUES_CHANGED」，故这里**不需要**再调 `commitValues()` / `notifyValuesChanged()`。
 *
 * @param data - 主数据集合
 * @param writesJson - Worker 回传的写回 JSON（`{"A1": "新值", ...}`，空串表示清空）
 * @param label - 撤销栈里的操作标签（用户在撤销菜单里看到的名字）
 * @returns {number} 实际写回的格子数
 */
export function applyWrites(data: DataCollection, writesJson: string, label: string): number {
    const raw = JSON.parse(writesJson) as Record<string, string>;
    const names = Object.keys(raw);
    if (names.length === 0) return 0;

    // 【合并写入重定向】脚本写「被合并覆盖的位置」（D3:F5 里的 E4）时一律落到锚点 D3，
    // 与 UI 侧完全同口径（详见 DataCollection#ensureWritableCell）：
    // 那些位置在视觉上根本不存在（被 drawMergedCellCover 的遮挡层盖住），按名新建
    // 只会往 values 里塞一条永远看不见、却能被「查找/替换」扫到的垃圾记录。
    // 锚点不在 values 时才新建，且**只新建锚点**（ensureWritableCell 内部保证）。
    // 多个被覆盖位置落到同一锚点时，按脚本的书写顺序取最后一个（后者覆盖前者）。
    const writes = new Map<string, string>();
    for (const name of names) {
        writes.set(data.resolveWriteTargetName(name), raw[name]);
    }

    const cells: Cell[] = [];
    writes.forEach((_value, targetName) => {
        cells.push(data.ensureWritableCell(targetName));
    });

    data.runWithCellsUndo(label, cells, () => {
        for (const cell of cells) {
            const value = writes.get(cell.cell);
            if (value === undefined || value === '') cell.clearValue();
            else cell.addValue(value);
        }
    });

    return cells.length;
}
