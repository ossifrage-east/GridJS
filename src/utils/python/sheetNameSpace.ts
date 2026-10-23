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
import { stripExtension, toArchiveKey } from '../fileName';
import { getCellDisplayText } from '../numberFormat';

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

「脚本」页签里的脚本可以读写**当前工作表**，也能用 open_grid() 打开工作区里的**其它表格**。

一、当前表格默认叫 sheet，三种写法任选（可混用）

    print(sheet.A1)              # 文本（等价于 get_cell("A1")）
    print(sheet.A1.value)        # 同上
    print(sheet.A1.color)        # 字体颜色
    print(sheet.A1.bgColor)      # 背景色
    print(sheet.A1.colspan)      # 合并跨度
    sheet.A1 = 56                # 写值
    sheet.A1.bgColor = "#fff"    # 写背景色

    print(sheet.cells(4, 3))         # （列号, 行号）→ D3
    print(sheet.cells("D", 3))       # 列字母也可以
    print(sheet.cells("D3"))         # 直接给单元格名
    sheet.cells(4, 3) = "12"         # 写值
    sheet.cells(4, 3).colspan = 3    # 设合并跨度

二、函数式 / 变量名式（老写法仍然可用）

    print(sheet_name, row_count, col_count)
    print(sheet.maxCol, sheet.maxRow)   # 有内容的最大列号 / 行号（1 起始，空表 0）
    print(files)                        # 工作区里所有表格文件（不含文件夹），可直接 open_grid()
    get_cell("A1") / set_cell("A1", "Hello") / clear_cell("A1")
    get_range("A1:C3") / set_range("A1:C3", [["a1","b1"],["a2","b2"]])
    A1 = 1            # 等价于 set_cell("A1", 1)（列字母须大写，模块级与函数体内都成立）
    x = A1            # 等价于 x = get_cell("A1")

三、打开其它表格（同一个工作区里的文件）

    other = open_grid("明细.xlsx")        # 相对当前文件所在目录（可用 . / .. / 子目录）
    other = open_grid("../汇总.xlsx")     # 上一级目录
    other = open_grid("/数据/明细")       # 以 / 开头 = 从根算起，扩展名可省
    print(other.B2, other.B2.bgColor)
    other.B2 = "改这一格"                 # 写回的是**那个文件**的存档，不影响当前表
    other.cells(2, 2).bold = True
    print(other["B2"], other.get_range("B2:C2"), other.title)

    # 刻意不叫 open：内置 open() 是读普通文件用的，重名会把文本读写一起截走。
    # 读文件照旧 open("a.txt")，读表格用 open_grid("a.xlsx")，两者互不干扰。

四、可读写的属性（读写同名；赋值即写回表格）

    value / text      文本内容
    color             字体颜色        bgColor / bg     背景色
    font              字体族          size             字号
    bold              加粗            italic           斜体
    underline         下划线          strike           删除线
    align             水平对齐        valign           垂直对齐
    wrap              自动换行        numFmt           数字格式码
    border            边框颜色        borderTop / borderBottom / borderLeft / borderRight  边框粗细
    colspan / rowspan 合并跨度

五、读到的内容是「表格最终显示的样子」

    print(sheet.A1)        # 与画布完全一致：按单元格的数字格式渲染
    sheet.A1.value         # 同上（最终显示文本）
    sheet.A1.display       # 同上，显式写法
    sheet.A1.raw           # 原始文本（chars 拼接，未格式化）
    sheet.A1.number        # 数值（日期格子给 Excel 序列号）
    sheet.A1.datetime      # datetime 对象（数字按序列号换算，文本按常见日期写法解析）
    sheet.A1.date          # date 对象
    sheet.A1.numFmt        # 数字格式码（如 yyyy/mm/dd）

    日期 / 时间 / 货币 / 百分比 / 科学计数 / 千分位都走同一套格式化，
    所以 print 出来的就是用户在表格里看到的那串字符：

        # 单元格存 45646、格式 yyyy/mm/dd
        print(sheet.A1)              # 2024/12/20
        print(sheet.A1.datetime)     # 2024-12-20 00:00:00
        print(sheet.A1 + 1)          # 45647（算术用原始序列号）
        print(sheet.A1 == "2024/12/20", sheet.A1 == 45646)   # True True

    写日期时可以直接给 datetime / date（字符串也行），会自动换算成 Excel 序列号，
    显示成什么由单元格自己的日期格式决定：

        import datetime
        sheet.A1.datetime = datetime.datetime(2026, 10, 23)
        sheet.A2.date = "2026-10-23"        # 字符串也认

注意：本次运行**写进去**的值，脚本内再读回来是写入的原样文本 —— 格式化由宿主完成，
运行结束后表格才会显示成格式化后的样子。

写「被合并覆盖的位置」（如 D3:F5 里的 E4）与界面同口径：一律落到锚点 D3。
脚本运行结束后，所有改动由宿主一次性回写（当前表进撤销栈，其它表格直接落存档）。
"""

import ast as _ast
import datetime as _datetime
import json as _json
import re as _re

# 已可用的表格：规范化路径 -> _Sheet（同一个对象会有多个别名，见 _register）
_OPENED = {}

# 当前工作表对象（_load 里创建；open_grid() 打开的表格是另外的 _Sheet 对象）
sheet = None
# 当前工作表名（脚本的存储键）；_SheetRef 是 str 子类，既能当表名也能 sheet_name.A1
sheet_name = ""
# 当前表格行数
row_count = 0
# 当前表格列数
col_count = 0
# 工作区里所有**表格**文件（根相对路径，不含文件夹）；可直接 open_grid(files[0])
files = []

_CELL_NAME = _re.compile(r"^([A-Za-z]{1,3})([0-9]{1,7})$")
# 纯列字母（cells 的第一个参数允许写成 "D"）
_COL_LETTERS = _re.compile(r"^[A-Za-z]{1,3}$")
# 变量名形态的单元格标识：**列字母大写** + 行号（A1、C10、AA100）。
# 刻意只认大写：x1 / b2 / a1 这类小写名在 Python 里太常见，
# 若也当成单元格名，普通脚本会被静默改成「往表里写值」，代价远大于收益。
_VAR_NAME = _re.compile(r"^[A-Z]{1,3}[0-9]{1,7}$")

# 用户侧属性名 -> 宿主 Cell 的规范化字段名（宿主按这个名字回写）
_FIELDS = {
    "value": "value",
    "text": "value",
    "color": "fontColor",
    "fontColor": "fontColor",
    "bgColor": "backgroundColor",
    "bg": "backgroundColor",
    "backgroundColor": "backgroundColor",
    "font": "fontFamily",
    "fontFamily": "fontFamily",
    "size": "fontSize",
    "fontSize": "fontSize",
    "bold": "fontWeight",
    "fontWeight": "fontWeight",
    "italic": "fontStyle",
    "fontStyle": "fontStyle",
    "underline": "underline",
    "strike": "strikethrough",
    "strikethrough": "strikethrough",
    "align": "textAlign",
    "textAlign": "textAlign",
    "valign": "alignItems",
    "alignItems": "alignItems",
    "wrap": "wrap",
    "numFmt": "numberFormat",
    "numberFormat": "numberFormat",
    "border": "borderColor",
    "borderColor": "borderColor",
    "borderTop": "borderTopWidth",
    "borderTopWidth": "borderTopWidth",
    "borderBottom": "borderBottomWidth",
    "borderBottomWidth": "borderBottomWidth",
    "borderLeft": "borderLeftWidth",
    "borderLeftWidth": "borderLeftWidth",
    "borderRight": "borderRightWidth",
    "borderRightWidth": "borderRightWidth",
    "letterSpacing": "letterSpacing",
    "lineSpacing": "lineSpacing",
    "colspan": "colspan",
    "rowspan": "rowspan",
}

# _Sheet 自己的字段（其余属性访问都当成单元格名）
_SHEET_ATTRS = ("key", "title", "rows", "cols", "snapshot", "styles", "writes", "vals", "display_map")


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
    cleaned = str(name).strip().replace("$", "")
    m = _CELL_NAME.match(cleaned)
    if not m:
        raise ValueError("不是合法的单元格名：%s" % (name,))
    return _col_index(m.group(1)), int(m.group(2))


def _name(col, row):
    """由列号、行号生成规范单元格名（如 A1）。"""
    return "%s%d" % (_col_name(col), row)


def _parse_range(a1):
    """解析 "A1:C3" 或单格 "A1"，返回按输入顺序的 (c1, r1, c2, r2)。"""
    parts = [p for p in str(a1).strip().split(":") if p.strip()]
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


def _canonical(name):
    """把任意写法的单元格名（a1 / $D$3 / d3）规范成 A1 形态。"""
    col, row = _parse(name)
    return _name(col, row)


def _to_number(text):
    """文本转数字：能转就转（int 优先），否则返回 None。"""
    if text is None or text == "":
        return None
    try:
        return int(text)
    except (TypeError, ValueError):
        pass
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def _to_datetime(text):
    """文本 -> datetime：纯数字按 Excel 序列号，其余按常见日期写法解析；都不行返回 None。

    序列号基准是 1899-12-30（Excel 的 1 = 1900-01-01）。Excel 把 1900 年当闰年，
    故序列号 < 60 那一段的基准要晚一天，否则 1900 年前两个月会差一天。
    """
    trimmed = (text or "").strip()
    if not trimmed:
        return None
    num = _to_number(trimmed)
    if num is not None:
        base = _datetime.datetime(1899, 12, 30)
        if num < 60:
            base = _datetime.datetime(1899, 12, 31)
        try:
            return base + _datetime.timedelta(days=float(num))
        except (OverflowError, ValueError):
            return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d"):
        try:
            return _datetime.datetime.strptime(trimmed, fmt)
        except ValueError:
            continue
    return None


def _to_serial(value):
    """datetime / date / 日期字符串 -> Excel 序列号；其它类型返回 None。"""
    if isinstance(value, str):
        parsed = _to_datetime(value)
        if parsed is None:
            return None
        value = parsed
    if isinstance(value, _datetime.datetime):
        delta = value - _datetime.datetime(1899, 12, 30)
    elif isinstance(value, _datetime.date):
        delta = _datetime.datetime(value.year, value.month, value.day) - _datetime.datetime(1899, 12, 30)
    else:
        return None
    return delta.days + delta.seconds / 86400.0


def _number_text(value):
    """数字 -> 写入用的文本（整数值不带 .0，小数去掉尾随零）。"""
    if value == int(value):
        return str(int(value))
    return ("%f" % value).rstrip("0").rstrip(".")


def _sheet_key(path):
    """路径规范化：去空白、反斜杠换正斜杠、转小写（Windows 路径不区分大小写）。"""
    return str(path).strip().replace(chr(92), "/").lower()


def _register(alias, target):
    """登记一个别名指向同一张表。"""
    if not alias:
        return
    _OPENED[_sheet_key(alias)] = target


class _Cell:
    """单元格对象：文本 + 样式 + 跨度的读写代理。

    读：print(sheet.A1) 打印文本（__str__），属性读样式；
    写：sheet.A1 = 56 / sheet.A1.bgColor = "#fff" 都直接记进所属表的 writes，
        脚本结束后由宿主一次性回写（当前表进撤销栈，其它表格落存档）。

    刻意不做成「延迟写」对象：所有读写都实时落在表对象上，故同一脚本里
    「写完再读」拿到的一定是新值。
    """

    __slots__ = ("sheet", "name")

    def __init__(self, sheet_obj, name):
        object.__setattr__(self, "sheet", sheet_obj)
        object.__setattr__(self, "name", name)

    # —— 属性读写 ——

    def __getattr__(self, attr):
        # —— 内容视图（只读）——
        # raw：chars 拼接的原文（未格式化）；display：与画布一致的最终显示文本
        if attr == "raw":
            return self.sheet.text(self.name)
        if attr == "display":
            return self.sheet.display(self.name)
        if attr == "number":
            # 数值：原文能转数字就给数字（日期单元格给的是 Excel 序列号）
            return _to_number(self.sheet.text(self.name))
        if attr in ("datetime", "date"):
            moment = _to_datetime(self.sheet.text(self.name))
            if moment is None:
                return None
            return moment if attr == "datetime" else moment.date()
        field = _FIELDS.get(attr)
        if field is None:
            raise AttributeError(
                "单元格 %s 没有属性 %s（可用：value/raw/display/number/datetime/date/"
                "color/bgColor/font/size/bold/italic/underline/strike/align/valign/wrap/"
                "numFmt/border/colspan/rowspan）" % (self.name, attr))
        if field == "value":
            # value 与 print() 一致：给**最终显示文本**（日期格式化后就是日期串）
            return self.sheet.display(self.name)
        return self.sheet.style(self.name, field)

    def __setattr__(self, attr, value):
        if attr in ("raw", "display", "number"):
            raise AttributeError("%s 是只读属性（它是格子内容的视图，不是样式）" % (attr,))
        if attr in ("datetime", "date"):
            # 写日期：转成 Excel 序列号写进格子，单元格设了日期格式就会显示成日期
            serial = _to_serial(value)
            if serial is None:
                raise ValueError("%s 需要 datetime / date 对象，收到：%r" % (attr, value))
            self.sheet.set_value(self.name, _number_text(serial))
            return
        field = _FIELDS.get(attr)
        if field is None:
            raise AttributeError("单元格 %s 没有属性 %s（赋值前请检查拼写）" % (self.name, attr))
        if field == "value":
            self.sheet.set_value(self.name, value)
        else:
            self.sheet.write(self.name, field, value)

    # —— 文本协议（让单元格像字符串一样参与运算）——

    def __str__(self):
        return self.sheet.display(self.name)

    def __repr__(self):
        return "%s = %r" % (self.name, self.sheet.display(self.name))

    def __len__(self):
        return len(self.sheet.display(self.name))

    def __bool__(self):
        text = self.sheet.text(self.name)
        num = _to_number(text)
        if num is not None:
            return num != 0
        return text != ""

    def __int__(self):
        return int(_to_number(self.sheet.text(self.name)) or 0)

    def __float__(self):
        return float(_to_number(self.sheet.text(self.name)) or 0.0)

    def __eq__(self, other):
        if isinstance(other, _Cell):
            return self.sheet.display(self.name) == other.sheet.display(other.name)
        if isinstance(other, (int, float)) and not isinstance(other, bool):
            num = _to_number(self.sheet.text(self.name))
            return num == other if num is not None else False
        return self.sheet.display(self.name) == other

    def __ne__(self, other):
        return not self.__eq__(other)

    def __add__(self, other):
        return _numeric(self) + other

    def __radd__(self, other):
        return other + _numeric(self)

    def __sub__(self, other):
        return _numeric(self) - other

    def __rsub__(self, other):
        return other - _numeric(self)

    def __mul__(self, other):
        return _numeric(self) * other

    def __rmul__(self, other):
        return other * _numeric(self)

    def __truediv__(self, other):
        return _numeric(self) / other

    def __rtruediv__(self, other):
        return other / _numeric(self)


def _numeric(value):
    """参与算术时把单元格换成「数字（能转的话）或文本」。"""
    if isinstance(value, _Cell):
        text = value.sheet.text(value.name)
        num = _to_number(text)
        return text if num is None else num
    return value


def _as_cell_name(attr):
    """属性名是否长得像单元格名（A1 / a1 / AA100）；不是则返回 None。"""
    if not isinstance(attr, str):
        return None
    if _CELL_NAME.match(attr):
        return _canonical(attr)
    return None


class _Sheet:
    """一张工作表：属性 / 下标 / cells() 三种方式都能定位单元格。

    内部持有**本表**的文本快照、样式快照与本次运行的写入，
    故 open_grid() 出来的多张表互不干扰，_dump 时按表分别回传。
    """

    def __init__(self, key, title, rows, cols, snapshot=None, styles=None, max_col=0, max_row=0,
                 display_map=None):
        object.__setattr__(self, "key", key)
        object.__setattr__(self, "title", title)
        object.__setattr__(self, "rows", rows)
        object.__setattr__(self, "cols", cols)
        object.__setattr__(self, "_max_col", max_col or 0)
        object.__setattr__(self, "_max_row", max_row or 0)
        # 宿主算好的「最终显示文本」（只含与原文不同的格子）；写入过的格子以写入为准
        object.__setattr__(self, "display_map", display_map if display_map else {})
        object.__setattr__(self, "snapshot", snapshot if snapshot else {})
        object.__setattr__(self, "styles", styles if styles else {})
        object.__setattr__(self, "writes", {})
        object.__setattr__(self, "vals", {})

    # —— 内部读写（_Cell 都走这里）——

    def text(self, name):
        """取文本：本次写入优先，其次快照；空格子为空串。"""
        w = self.writes.get(name)
        if w is not None and "value" in w:
            return w["value"]
        return self.snapshot.get(name, "")

    def display(self, name):
        """取**最终显示文本**（与画布一致：按 numberFormat 格式化过）。

        本次运行写入过的格子返回写入原文 —— 格式化在宿主侧做，脚本内拿不到
        写完之后再格式化的结果（运行结束后表格会按格式显示）。
        """
        w = self.writes.get(name)
        if w is not None and "value" in w:
            return w["value"]
        shown = self.display_map.get(name)
        if shown is not None:
            return shown
        return self.snapshot.get(name, "")

    def style(self, name, field):
        """取样式字段：本次写入优先，其次样式快照；没有则 None。"""
        w = self.writes.get(name)
        if w is not None and field in w:
            return w[field]
        styles = self.styles.get(name)
        if styles is not None:
            return styles.get(field)
        return None

    def write(self, name, field, value):
        """记一笔写入（同一格多次写会合并到同一个字典）。"""
        w = self.writes.get(name)
        if w is None:
            w = {}
            self.writes[name] = w
        if field == "value":
            w["value"] = "" if value is None else str(value)
            return
        w[field] = value

    def set_value(self, name, value):
        """写值：原始 Python 值留一份（让 A1 + 1 按数值算），表层写字符串。"""
        self.vals[name] = value
        self.write(name, "value", value)

    # —— 内容范围（只读，由格子算出来；本次运行写入的格子也计入）——

    @property
    def maxCol(self):
        """有内容的最大列号（1 起始；空表 0）。"""
        best = self._max_col
        for name in self.writes:
            col = _parse(name)[0]
            if col > best:
                best = col
        return best

    @property
    def maxRow(self):
        """有内容的最大行号（1 起始；空表 0）。"""
        best = self._max_row
        for name in self.writes:
            row = _parse(name)[1]
            if row > best:
                best = row
        return best

    # —— 对象式 ——

    def __getattr__(self, attr):
        name = _as_cell_name(attr)
        if name is None:
            raise AttributeError(
                "表格 %r 没有属性 %s（单元格请写成 sheet.A1；打开表格用 open_grid(...)）"
                % (self.title, attr))
        return _Cell(self, name)

    def __setattr__(self, attr, value):
        if attr in ("maxCol", "maxRow"):
            raise AttributeError("%s 是只读属性（按表格内容算出来的）" % (attr,))
        if attr in _SHEET_ATTRS or attr.startswith("_"):
            object.__setattr__(self, attr, value)
            return
        name = _as_cell_name(attr)
        if name is None:
            raise AttributeError("表格没有属性 %s（写值请写成 sheet.A1 = 1）" % (attr,))
        self.set_value(name, value)

    # —— 下标式 / cells() ——

    def _key_name(self, key):
        if isinstance(key, tuple):
            if len(key) != 2:
                raise ValueError("单元格下标需要（列号, 行号）两个值，例如 sheet[4, 3]")
            return _name(int(key[0]), int(key[1]))
        if isinstance(key, int):
            raise ValueError("单元格下标需要（列号, 行号）两个值，例如 sheet[4, 3]")
        return _canonical(key)

    def __getitem__(self, key):
        if isinstance(key, str) and ":" in key:
            return self.get_range(key)
        name = self._key_name(key)
        if name in self.vals:
            return self.vals[name]
        return self.text(name)

    def __setitem__(self, key, value):
        self.set_value(self._key_name(key), value)

    def __delitem__(self, key):
        name = self._key_name(key)
        self.vals.pop(name, None)
        self.write(name, "value", "")

    def cells(self, col, row=None):
        """按（列号, 行号）取单元格对象；也接受列字母或单元格名。

        sheet.cells(4, 3) 与 sheet.cells("D", 3) 与 sheet.cells("D3") 都指向 D3。
        赋值的写法 sheet.cells(4, 3) = "12" 由 _rewrite_cells_assign 在编译期
        改写成 sheet[4, 3] = "12"（Python 语法本身不允许给函数调用赋值）。
        """
        if row is None:
            return _Cell(self, self._key_name(col))
        if isinstance(col, str):
            letters = col.strip().replace("$", "")
            if not _COL_LETTERS.match(letters):
                raise ValueError("列号需要字母（如 D）或数字（如 4），收到：%r" % (col,))
            col = _col_index(letters)
        return _Cell(self, _name(int(col), int(row)))

    def get_range(self, a1):
        """取 "A1:C3" 区域的文本，返回二维列表（行 × 列，越界补空串）。"""
        c1, r1, c2, r2 = _parse_range(a1)
        out = []
        for row in range(min(r1, r2), max(r1, r2) + 1):
            out.append([self.text(_name(col, row)) for col in range(min(c1, c2), max(c1, c2) + 1)])
        return out

    def __repr__(self):
        return "<表格 %r（%d 行 × %d 列）>" % (self.title, self.rows, self.cols)

    def __str__(self):
        return self.title


class _SheetRef(str):
    """表名的字符串子类：既当表名打印，也能 sheet_name.A1 这样取单元格。"""

    def __new__(cls, sheet_obj):
        obj = str.__new__(cls, sheet_obj.title)
        object.__setattr__(obj, "_sheet", sheet_obj)
        return obj

    def __getattr__(self, attr):
        if attr.startswith("_"):
            raise AttributeError(attr)
        name = _as_cell_name(attr)
        if name is None:
            raise AttributeError("表名 %r 没有属性 %s" % (str(self), attr))
        return _Cell(object.__getattribute__(self, "_sheet"), name)

    def __setattr__(self, attr, value):
        if attr.startswith("_"):
            object.__setattr__(self, attr, value)
            return
        name = _as_cell_name(attr)
        if name is None:
            raise AttributeError("表名 %r 没有属性 %s" % (str(self), attr))
        object.__getattribute__(self, "_sheet").set_value(name, value)


class _Cells:
    """单元格命名空间：脚本里 A1 = 1 这类写法最终都落到这里（内部接口）。

    ## 为什么不直接拦截用户脚本的 globals

    A1 = 1 在**模块级**编译为 STORE_NAME、在**函数体内**编译为 STORE_FAST。
    后者直接写进帧对象的局部数组，任何字典子类的 __setitem__ 都拦不到；函数内
    再加 global 声明也无效（STORE_GLOBAL 走的是 C 层 PyDict_SetItem，同样绕过
    层钩子）。实测确认过这一点，故改为在编译前用 AST 改写（见 _CellRewrite）。

    该对象恒指**当前工作表**（_load 里绑定），open_grid() 出来的表格请用 .cells()。
    """

    def __init__(self, target):
        self.sheet = target

    def __getitem__(self, key):
        target = self.sheet
        name = target._key_name(key)
        if name in target.vals:
            return target.vals[name]
        return target.text(name)

    def __setitem__(self, key, value):
        target = self.sheet
        target.set_value(target._key_name(key), value)

    def __delitem__(self, key):
        target = self.sheet
        name = target._key_name(key)
        target.vals.pop(name, None)
        target.write(name, "value", "")


__cells__ = _Cells(None)


def open_grid(path):
    """打开工作区里的另一张表格，返回可读写的表格对象。

    ## 路径写法（相对路径以**当前文件所在目录**为基准）

        open_grid("明细.xlsx")          # 与当前文件同目录
        open_grid("./明细.xlsx")        # 同上（. 就是当前目录）
        open_grid("../汇总.xlsx")       # 上一级目录
        open_grid("../../归档/旧表.xlsx")  # 上两级再进子目录
        open_grid("/数据/明细.xlsx")    # 以 / 开头 = 从根算起
        open_grid("数据/明细")          # 省略扩展名也行（唯一匹配时）

    扩展名可省，大小写不敏感；文件树里任何有内容的表格都能打开。
    .. 越过根时停在根，不会跑到工作区外面。

    例：

        other = open_grid("数据/明细.xlsx")
        print(other.B2, other.B2.bgColor)
        other.B2 = "改这一格"        # 只改那个文件的存档，不影响当前表

    ## 为什么叫 open_grid 而不是 open

    内置 open() 是用来读普通文件的（Pyodide 里操作虚拟文件系统），
    重名会把脚本里的文本读写一起截走。故本模块**不遮蔽内置 open**，
    读文件照旧用 open()，读表格用 open_grid()。

    ## 只能打开**脚本里写成字符串字面量**的路径

    宿主在运行前静态扫出 open_grid("...") 的字面量路径并预加载，
    故 open_grid(变量) 拿不到数据 —— 那种写法会直接抛 ValueError。
    """
    found = _find_sheet(path)
    if found is None:
        names = sorted(set([obj.title for obj in _OPENED.values()]))
        raise ValueError(
            "打不开这张表格：%r。只支持脚本里写成字符串字面量的路径（如 open_grid('数据/明细.xlsx')），"
            "动态路径（open_grid(变量)）扫不到；当前可用的表格：%s"
            % (path, "、".join(names) if names else "无"))
    return found


def _find_sheet(path):
    """在已加载的表格里找 path（支持 _Sheet 自身的幂等返回）。"""
    if isinstance(path, _Sheet):
        return path
    return _OPENED.get(_sheet_key(path))


# 旧写法 Cells(4, 3) 作为 sheet.cells(4, 3) 的别名保留
def Cells(col, row=None):
    """按（列号, 行号）取**当前表**的单元格对象（等价于 sheet.cells(...)）。"""
    return sheet.cells(col, row)


def get_cell(name):
    """取单元格文本；空格子返回空串，本次已写入的值优先于快照。"""
    return sheet.text(_canonical(name))


def set_cell(name, value):
    """设置单元格文本；value 为 None 时等价清空；其余按 str(value) 写入。"""
    sheet.set_value(_canonical(name), value)


def clear_cell(name):
    """清空单元格。"""
    sheet.set_value(_canonical(name), "")


def get_range(a1):
    """取 "A1:C3" 区域的文本，返回二维列表（行 × 列，越界补空串）。"""
    return sheet.get_range(a1)


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
            sheet.set_value(_name(c1 + j, r1 + i), value)


def iter_cells():
    """迭代**当前表**所有非空格子，产出 (名称, 文本)；含本次写入的改动。"""
    keys = set(sheet.snapshot)
    keys.update(sheet.writes)
    for key in sorted(keys, key=_key_order):
        text = sheet.text(key)
        if text:
            yield key, text


class _CellRewrite(_ast.NodeTransformer):
    """把「单元格名形态」的标识符改写成 __cells__["A1"]（内部接口）。

    改写成下标后，A1 = 1 / x = A1 / A1 += 1 / A1, B1 = 1, 2 /
    for A1 in ... / del A1 全都自动成立 —— 下标本身就是合法的赋值目标。

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


# 「以 Cells(...) 开头的赋值语句」：行首缩进 + Cells(参数) + 赋值号
# 参数里不允许出现括号（覆盖 Cells(4, 3) / Cells("D", 3) / Cells("D3") 这几种形态）
_CELLS_ASSIGN = _re.compile(r"^([^A-Za-z0-9_]*?)Cells\\(([^()]*)\\)[ ]*([+*/-]?=)(?!=)")
# 「x.cells(...) 赋值」：把 .cells(...) 换成 [...]（任意对象上的 cells 方法）
_METHOD_ASSIGN = _re.compile(r"\\.cells\\(([^()]*)\\)[ ]*([+*/-]?=)(?!=)")


def _cell_target(args):
    """把 cells(...) 的参数拼成下标表达式（1 个参数当名字，2 个当列,行）。"""
    return "[" + args + "]"


def _rewrite_one_line(line):
    """改写一行里的 Cells(...) / x.cells(...) 赋值（内部接口）。"""
    stripped = line.lstrip()
    if stripped.startswith("#"):
        return line
    # 引号出现在匹配位置之前说明这是字符串内容，不是代码 —— 不碰（宁可报语法错，
    # 也不能把 print("a.cells(1) = b") 悄悄改成 print("a[1] = b")）
    quote_at = min([i for i in (line.find('"'), line.find("'")) if i >= 0] or [len(line) + 1])

    m = _CELLS_ASSIGN.match(line)
    if m is not None and m.start() < quote_at:
        line = m.group(1) + "__cells__" + _cell_target(m.group(2)) + " " + m.group(3) + line[m.end():]

    m2 = _METHOD_ASSIGN.search(line)
    if m2 is not None and m2.start() < quote_at:
        line = line[:m2.start()] + _cell_target(m2.group(1)) + " " + m2.group(2) + line[m2.end():]
    return line


def _rewrite_cells_assign(source):
    """逐行改写「给 cells(...) 赋值」的写法（内部接口）。

    Python **语法本身**不允许给函数调用赋值：sheet.cells(4, 3) = "12" 在
    ast.parse 阶段就报 SyntaxError（cannot assign to function call here），
    所以 AST 改写根本轮不到 —— 只能在源码层面先换成下标赋值：

        sheet.cells(4, 3) = "12"    ->    sheet[4, 3] = "12"
        Cells(4, 3) = "12"          ->    __cells__[4, 3] = "12"

    只处理**单行、参数里不含括号、匹配位置在第一个引号之前**的形态；
    其余保持原样，由 Python 自己报语法错误（不会静默改成别的意思）。
    逐行替换不改变行数，traceback 的行号仍与用户脚本一致。
    """
    return chr(10).join([_rewrite_one_line(line) for line in source.split(chr(10))])


def _compile(source, filename):
    """编译用户脚本：先做源码级改写 + AST 改写，再交给内置 compile（内部接口）。

    两道改写缺一不可：
    - _rewrite_cells_assign：给 cells(...) 赋值（语法非法，必须先改源码）；
    - _CellRewrite：A1 = 1 这类单元格名标识符（改写成下标）。
    """
    text = _rewrite_cells_assign(source)
    tree = _CellRewrite().visit(_ast.parse(text, filename, "exec"))
    _ast.fix_missing_locations(tree)
    return compile(tree, filename, "exec")


def _load(snapshot_json, styles_json, opened_json, name, rows, cols, meta_json=""):
    """宿主在运行脚本前调用：注入当前表、open_grid() 出来的表格与元信息（内部接口）。

    meta_json 带三样东西（都可缺省）：maxCol / maxRow（当前表有内容的最大行列号）、
    files（工作区里所有表格文件，根相对路径）。缺省时按空表 / 空清单处理。
    """
    global sheet, sheet_name, row_count, col_count, files, _OPENED
    _OPENED = {}
    meta = _json.loads(meta_json) if meta_json else {}
    files = list(meta.get("files") or [])
    current = _Sheet(name, name, rows, cols, _json.loads(snapshot_json), _json.loads(styles_json),
                     meta.get("maxCol") or 0, meta.get("maxRow") or 0, meta.get("display"))
    sheet = current
    _register(name, current)
    for item in _json.loads(opened_json) if opened_json else []:
        if item.get("current"):
            target = current
        else:
            target = _Sheet(item["key"], item.get("title") or item["key"], item.get("rows") or 0,
                            item.get("cols") or 0, item.get("snapshot"), item.get("styles"),
                            item.get("maxCol") or 0, item.get("maxRow") or 0, item.get("display"))
        _register(item["key"], target)
        for alias in item.get("aliases", []):
            _register(alias, target)
    __cells__.sheet = sheet
    sheet_name = _SheetRef(sheet)
    row_count = rows
    col_count = cols


def _dump():
    """宿主在运行脚本后调用：按表取出本次写入（内部接口）。

    形态：{"__sheets__": {"<表键>": {"A1": {"value": "x", ...}}}}
    只包含**确实写过**的表，空表不出现。
    """
    out = {}
    seen = set()
    for obj in _OPENED.values():
        if id(obj) in seen:
            continue
        seen.add(id(obj))
        if obj.writes:
            out[obj.key] = obj.writes
    return _json.dumps({"__sheets__": out})
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
    /** 表格列数（表头宽度，不是内容范围） */
    colCount: number;
    /** 有内容的最大列号（1 起始；空表 0，见 DataCollection#getContentBounds） */
    maxCol: number;
    /** 有内容的最大行号（1 起始；空表 0） */
    maxRow: number;
    /** 非空格子（`[单元格名, 文本]`） */
    cells: [string, string][];
    /** 发给 Python 的文本载荷（`{"A1": "文本"}`） */
    json: string;
    /** 发给 Python 的样式载荷（`{"A1": {"fontColor": "#fff"}}`，只含有样式的格子） */
    stylesJson: string;
    /**
     * 最终显示文本（按 numberFormat 格式化后），**只收与原文不同的格子**
     *
     * 与 {@link displayJson} 内容相同，区别只是一个对象、一个字符串：
     * 对象形态给 {@link buildScriptMeta} 塞进 metaJson，字符串形态供其它调用方直传。
     */
    display: Record<string, string>;
    /** {@link display} 的 JSON 形态 */
    displayJson: string;
}

/**
 * Python 侧可读的样式字段（与 `PY_NAMESPACE_SOURCE` 里 `_FIELDS` 的右值一一对应）
 *
 * 一律用 `Cell` 的**规范化属性名**：Python 侧的别名（`color` / `bgColor` / `bold` …）
 * 在本地翻译成这些名字，宿主回写时按同一套名字赋值，两端的表必须同时改。
 */
const SCRIPT_STYLE_FIELDS: readonly string[] = [
    'fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'fontColor',
    'underline', 'strikethrough',
    'backgroundColor', 'borderColor',
    'borderTopWidth', 'borderBottomWidth', 'borderLeftWidth', 'borderRightWidth',
    'textAlign', 'alignItems', 'wrap', 'letterSpacing', 'lineSpacing',
    'numberFormat', 'colspan', 'rowspan',
];

/**
 * 收集一个格子上「脚本可读」的样式（没有样式时返回 null）
 *
 * 注意字体字段是 `Cell` 上的**聚合投影**（`fontFamily` 读 `chars[0]`）：富文本多段
 * 样式只能读到首字符那一段。写回则是单元格级（setter 写全部字符），与界面工具栏
 * 的粒度一致 —— 脚本层面不做「逐字符改样式」。
 * @param cell - 目标单元格
 * @returns {Record<string, unknown> | null} 样式字典；无任何样式时为 null
 */
function collectCellStyle(cell: Cell): Record<string, unknown> | null {
    let out: Record<string, unknown> | null = null;
    const raw = cell as unknown as Record<string, unknown>;
    for (let i = 0; i < SCRIPT_STYLE_FIELDS.length; i++) {
        const key = SCRIPT_STYLE_FIELDS[i];
        const value = raw[key];
        // false / 0 是有效值（如 bold=False、borderTop=0），只有「没设过」才跳过
        if (value === undefined || value === null || value === '') continue;
        if (out === null) out = {};
        out[key] = value;
    }
    return out;
}

/**
 * 导出当前工作表的快照（运行时送给 Python 的输入）
 *
 * 文本只带**非空**格子：空串在 Python 侧与「不存在」等价（`get_cell` 返回 `""`），
 * 而百万行大表里空格子往往占绝大多数，带上它们只是徒增序列化成本。
 * 样式另走一份字典，且只收「确实有样式」的格子 —— 纯数据表这一份基本是空的，
 * 有格式的报表才会变大（每格几十字节）。
 * @param data - 主数据集合
 * @returns {SheetSnapshot} 快照
 */
export function buildSnapshot(data: DataCollection): SheetSnapshot {
    const raw = collectSnapshot(data);
    // 内容范围：与筛选下拉共用同一份索引（cells 里的文本 + 合并跨度，空格子不计）
    const bounds = data.getContentBounds();
    return {
        sheetName: data.activedSheetName ?? '',
        rowCount: data.rowHeaders.length,
        colCount: data.colHeaders.length,
        maxCol: bounds.maxCol,
        maxRow: bounds.maxRow,
        cells: raw.cells,
        json: JSON.stringify(raw.text),
        stylesJson: JSON.stringify(raw.styles),
        display: raw.display,
        displayJson: JSON.stringify(raw.display),
    };
}

/** 快照的「对象形态」（字符串形态见 {@link buildSnapshot}） */
interface RawSnapshot {
    /** 非空格子 `[名称, 文本]` */
    cells: [string, string][];
    /** 单元格名 → 原始文本（chars 拼接，未格式化） */
    text: Record<string, string>;
    /** 单元格名 → **最终显示文本**（按 numberFormat 格式化；只收与原文不同的格子） */
    display: Record<string, string>;
    /** 单元格名 → 样式（只含有样式的格子） */
    styles: Record<string, Record<string, unknown>>;
}

/**
 * 把一张表的可读内容收集成内存对象（{@link buildSnapshot} 与多表预加载共用）
 * @param data - 任意数据集合（当前表或一张临时装载的表）
 * @returns {RawSnapshot} 对象形态快照
 */
function collectSnapshot(data: DataCollection): RawSnapshot {
    const cells: [string, string][] = [];
    const text: Record<string, string> = {};
    const display: Record<string, string> = {};
    const styles: Record<string, Record<string, unknown>> = {};
    for (const cell of data.values) {
        const name = cell.cell;
        const value = cell.getText();
        if (value !== '') {
            cells.push([name, value]);
            text[name] = value;
            // 最终显示文本必须与画布完全一致，故**只**走渲染层这同一个函数
            // （日期 / 货币 / 百分比 / 科学计数都在里面）。无格式的格子它原样返回，
            // 故这里不会为它们产生条目 —— display 是稀疏的。
            const shown = getCellDisplayText(value, cell.numberFormat);
            if (shown !== value) display[name] = shown;
        }
        const style = collectCellStyle(cell);
        if (style !== null) styles[name] = style;
    }
    return { cells, text, display, styles };
}

/**
 * 把脚本的写回改动落到当前工作表上（纳入撤销栈）
 *
 * 只做「载荷归一化」：把字符串简写（`{"A1": "新值"}`）与字段字典
 * （`{"A1": {"value": "新值", "bgColor": "#fff"}}`）统一成字段字典，其余
 * 落点解析 / 样式赋值 / 合并跨度清理 / 撤销记账全部交给
 * {@link DataCollection.applyScriptWrites} —— 那些规则（合并重定向、覆盖格清理、
 * 索引失效）必须和界面操作共用同一份实现，否则脚本写出来的表会与手点出来的不一样。
 * @param data - 主数据集合
 * @param writesJson - Worker 回传的写回 JSON（值为字符串或字段字典）
 * @param label - 撤销栈里的操作标签（用户在撤销菜单里看到的名字）
 * @returns {number} 实际写回的格子数（被合并重定向折叠后）
 */
export function applyWrites(data: DataCollection, writesJson: string, label: string): number {
    const parsed = JSON.parse(writesJson) as Record<string, unknown>;
    // 多表载荷（`{"__sheets__": {"某键": {格子: 字段}}}`）：只取当前表那一份
    const sheets = parsed.__sheets__ as Record<string, Record<string, unknown>> | undefined;
    if (sheets !== undefined) {
        const mine = sheets[data.activedSheetName];
        return mine === undefined ? 0 : data.applyScriptWrites(label, normalizePayload(mine));
    }
    return data.applyScriptWrites(label, normalizePayload(parsed as Record<string, unknown>));
}

/**
 * 把「单元格名 → 字符串 / 字段字典」的载荷统一成字段字典
 *
 * 两种形态都收：纯字符串（只写值，早期格式与手写调用）与字段字典
 * （`{"A1": {"value": "x", "bgColor": "#fff"}}`，脚本按属性写出来的形状）。
 * @param raw - 原始载荷
 * @returns {Record<string, Record<string, unknown>>} 字段字典
 */
function normalizePayload(raw: Record<string, unknown>): Record<string, Record<string, unknown>> {
    const payload: Record<string, Record<string, unknown>> = {};
    for (const name of Object.keys(raw)) {
        const entry = raw[name];
        payload[name] = typeof entry === 'string' ? { value: entry } : (entry as Record<string, unknown>);
    }
    return payload;
}

/**
 * 应用一次运行的全部写回：当前表进撤销栈，`open_grid()` 打开的其它表格直接落存档
 *
 * 为什么要分开：其它表格不在内存里（只是为这次运行临时装载的），写它等价于
 * 「改那个文件的存档」，既不能进当前表的撤销栈，也不该触碰当前表的数据。
 * @param data - 当前工作表
 * @param writesJson - Worker 回传的写回 JSON
 * @param label - 当前表撤销栈里的操作标签
 * @param scratch - 预加载的其它表格（键 = 存档键）；缺省时只处理当前表
 * @returns {Promise<{count: number, others: string[]}>} 写入格数 + 被改动的其它表格
 */
export async function applyWritesAll(
    data: DataCollection,
    writesJson: string,
    label: string,
    scratch?: Map<string, DataCollection>
): Promise<{ count: number; others: string[] }> {
    const parsed = JSON.parse(writesJson) as Record<string, unknown>;
    const sheets = parsed.__sheets__ as Record<string, Record<string, unknown>> | undefined;
    if (sheets === undefined) return { count: applyWrites(data, writesJson, label), others: [] };

    let count = 0;
    const others: string[] = [];
    for (const key of Object.keys(sheets)) {
        const payload = normalizePayload(sheets[key]);
        if (key === data.activedSheetName) {
            count += data.applyScriptWrites(label, payload);
            continue;
        }
        const target = scratch?.get(key);
        if (target === undefined) continue;      // 没预加载到的表格（理论上不会出现）
        count += target.applyScriptWrites(`运行 Python 脚本（${key}）`, payload);
        await target.syncValuesToDB();
        others.push(key);
    }
    return { count, others };
}

// ──────────────────────── open_grid() 多表支持 ────────────────────────

/**
 * 一张「脚本用 `open_grid()` 打开」的表格（随运行请求下发给 Python）
 *
 * {@link key} 是存档键，同时充当 Python 侧 `_dump()` 的分表键：回写时宿主
 * 按它判断「这份写入属于当前表还是别的文件」。
 */
export interface OpenedSheet {
    /** 存档键（= toArchiveKey(路径)） */
    key: string;
    /** 展示名（文件树里看到的路径；错误提示用） */
    title: string;
    /** 表格行数 */
    rowCount: number;
    /** 表格列数（表头宽度） */
    colCount: number;
    /** 有内容的最大列号（1 起始；空表 0） */
    maxCol: number;
    /** 有内容的最大行号（1 起始；空表 0） */
    maxRow: number;
    /** 用户可能写进 open_grid() 的其它写法（原路径 / 去扩展名 / 文件名） */
    aliases: string[];
    /** 与当前工作表是同一张表 → Python 侧直接复用当前对象，不重复加载 */
    current: boolean;
    /** 单元格名 → 文本 */
    text: Record<string, string>;
    /** 单元格名 → 样式 */
    styles: Record<string, Record<string, unknown>>;
    /** 单元格名 → 最终显示文本（按 numberFormat 格式化；只收与原文不同的格子） */
    display: Record<string, string>;
}

/** `open_grid("路径")` 的字面量形态（单参、纯字符串） */
const OPEN_GRID_CALL = /\bopen_grid\s*\(\s*(['"])([^'"\n]+)\1\s*\)/g;

/**
 * 从脚本源码里扫出 `open_grid("路径")` 的字面量参数
 *
 * 【为什么要在运行前静态扫一遍】Python 跑在 Worker 里，`open_grid()` 是**同步**调用，
 * 而读另一个文件的存档（IndexedDB）只能异步做 —— 同步 API 里等不了异步。
 * 因此改为「先扫出所有字面量路径 → 主线程预加载 → 再把数据一起下发给 Worker」，
 * 脚本里的 `open_grid()` 就只是查表，全程同步。
 *
 * 代价：**动态路径不支持**（`open_grid(p)` / `open_grid(a + ".xlsx")`）—— 那种写法扫不到，
 * 运行时 `open_grid()` 会给出明确报错。注释行、字符串内部的 `open_grid(` 也会被扫到，
 * 但顶多多加载一张表，不影响语义。
 * @param source - 用户脚本全文
 * @returns {string[]} 去重后的路径字面量
 */
export function extractOpenGridPaths(source: string): string[] {
    const out: string[] = [];
    const seen = new Set<string>();
    const lines = source.split('\n');
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.trimStart().startsWith('#')) continue;
        OPEN_GRID_CALL.lastIndex = 0;
        let hit = OPEN_GRID_CALL.exec(line);
        while (hit !== null) {
            const path = hit[2];
            if (!seen.has(path)) {
                seen.add(path);
                out.push(path);
            }
            hit = OPEN_GRID_CALL.exec(line);
        }
    }
    return out;
}

/**
 * 一张表在 Python 侧可以被 `open_grid()` 命中的所有写法
 * @param path - 用户写的路径
 * @param key - 存档键
 * @returns {string[]} 别名列表（首个是原路径）
 */
function buildAliases(input: string, key: string): string[] {
    const out: string[] = [];
    const push = (value: string): void => {
        if (value && out.indexOf(value) < 0) out.push(value);
    };
    // 用户写的原样字符串必须留在首位：Python 侧 open_grid("...") 是拿**字面量**查表的
    push(input);
    push(input.trim().replace(/\\/g, '/'));
    push(key);
    push(stripExtension(key));
    return out;
}

/**
 * 路径归一化：统一分隔符、折叠 `.` / `..`、去掉多余斜杠
 *
 * `..` 越过根时**钳制在根**（不会得到 `../x` 这种键）。
 * @param path - 任意写法的路径
 * @returns {string} 归一化后的相对路径（无前导 `/`）
 */
function normalizePath(path: string): string {
    const parts = (path ?? '').replace(/\\/g, '/').split('/');
    const out: string[] = [];
    for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        if (part === '' || part === '.') continue;
        if (part === '..') {
            out.pop();
            continue;
        }
        out.push(part);
    }
    return out.join('/');
}

/**
 * 取「某个存档键」所在目录（`数据/明细` → `数据`；根下文件 → 空串）
 * @param key - 存档键
 * @returns {string} 目录部分
 */
function dirOf(key: string): string {
    const at = (key ?? '').lastIndexOf('/');
    return at < 0 ? '' : key.slice(0, at);
}

/**
 * 把 `open_grid()` 里的路径解析成**存档键**
 *
 * ## 相对路径是相对「脚本当前文件所在目录」的
 *
 * `data.activedSheetName` 就是当前文件的存档键（≈ 路径），故它所在的目录就是基准目录：
 * - `"明细.xlsx"` → 同目录下的 `明细.xlsx`
 * - `"./明细.xlsx"` / `"../汇总.xlsx"` / `"子目录/明细.xlsx"` → 按常规相对路径折叠
 * - `"/数据/明细.xlsx"` → 从根算起
 *
 * ## 为什么还要「相对根」再试一次
 *
 * 早期版本把 `open_grid("数据/明细.xlsx")` 一律按**根**解析，已有脚本是这么写的；
 * 相对目录优先、根兜底，两种写法都不会突然打不开。
 *
 * ## 省略扩展名
 *
 * `"明细"` 也能命中 `数据/明细.xlsx` / `数据/明细`（`.grid` 的键没有后缀）：
 * 精确命中失败后，再按「去扩展名后相同」唯一匹配 —— **命中多张就当歧义、不猜**
 * （宁可报错让用户写清楚，也不能随便挑一张去写）。
 *
 * @param keys - 工作区里所有有存档的表格键（`IndexDB#getAllValues`）
 * @param currentKey - 当前文件的存档键（相对路径的基准目录取自它）
 * @param input - `open_grid()` 里写的路径
 * @returns {string | null} 命中的存档键；解析不到为 null
 */
export function resolveGridKey(keys: readonly string[], currentKey: string, input: string): string | null {
    const raw = (input ?? '').trim().replace(/\\/g, '/');
    if (!raw) return null;

    // 候选路径（按优先级）：相对当前文件目录 → 相对根
    const bases: string[] = [];
    const pushBase = (value: string): void => { if (value && bases.indexOf(value) < 0) bases.push(value); };
    if (raw.startsWith('/')) {
        pushBase(normalizePath(raw));
    } else {
        const dir = dirOf(currentKey);
        // 【顺序不能反】先拼再归一化：`..` 必须能吃掉基准目录里的上一级
        pushBase(normalizePath(dir ? `${dir}/${raw}` : raw));
        pushBase(normalizePath(raw));
    }

    // 索引为空（数据库读失败 / 空工作区）时退化：直接把候选当键试一次
    if (keys.length === 0) return bases.length > 0 ? toArchiveKey(bases[0]) : null;

    const lower = keys.map(key => key.toLowerCase());
    // ① 精确命中（不区分大小写；`.grid` 的键本身没有后缀，靠 ② 兜）
    for (let i = 0; i < bases.length; i++) {
        const at = lower.indexOf(toArchiveKey(bases[i]).toLowerCase());
        if (at >= 0) return keys[at];
    }
    // ② 省 / 换扩展名：去扩展名后唯一匹配
    for (let i = 0; i < bases.length; i++) {
        const bare = stripExtension(bases[i]).toLowerCase();
        let hit = -1;
        let count = 0;
        for (let j = 0; j < keys.length; j++) {
            if (stripExtension(keys[j]).toLowerCase() !== bare) continue;
            hit = j;
            count++;
        }
        if (count === 1) return keys[hit];
    }
    return null;
}

/**
 * 枚举工作区里所有「有存档的表格」键
 *
 * 这是「文件树里有哪些表」在数据层的等价物：文件树里的每个表格文件都有一条存档
 * （见 MEMORY 里的不变式），故按存档键枚举即可覆盖全部可打开的表。
 * 读失败时返回空数组 —— 解析会退化到「候选路径直接当键试」，不影响正常流程。
 * @param data - 当前工作表（借用它的 db 句柄）
 * @returns {Promise<string[]>} 存档键列表
 */
async function listArchiveKeys(data: DataCollection): Promise<string[]> {
    try {
        // 只取键，不读存档：大表工作区里 getAllValues() 会把几十万格的快照一起反序列化
        const names = await data.db.getAllSheetNames();
        return names.filter(name => typeof name === 'string' && name !== '');
    } catch (err) {
        return [];
    }
}

/**
 * 列出工作区里所有**表格**文件（不含文件夹），按路径排序
 *
 * 就是 Python 侧 `files` 列表的来源 —— 都是能直接喂给 `open_grid()` 的键（根相对路径）。
 * 只含表格：非表格文件（文档 / 演示文稿）没有表格存档，列出来也打不开。
 * @param data - 当前工作表（借用 db 句柄）
 * @returns {Promise<string[]>} 排序后的表格键
 */
export async function listGridFiles(data: DataCollection): Promise<string[]> {
    const keys = await listArchiveKeys(data);
    return keys.slice().sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
}

/** 随运行请求下发的「当前表元信息」 */
export interface ScriptMeta {
    /** 当前表有内容的最大列号（1 起始；空表 0） */
    maxCol: number;
    /** 当前表有内容的最大行号（1 起始；空表 0） */
    maxRow: number;
    /** 工作区里所有表格文件（根相对路径），即 Python 侧的 `files` */
    files: string[];
    /** 当前表的最终显示文本（只收与原文不同的格子），供 Python 侧 `print()` 用 */
    display: Record<string, string>;
}

/**
 * 组装当前表的元信息载荷
 * @param data - 当前工作表
 * @param files - {@link listGridFiles} 的结果
 * @returns {ScriptMeta} 元信息
 */
export function buildScriptMeta(
    data: DataCollection,
    files: string[],
    display: Record<string, string>
): ScriptMeta {
    const bounds = data.getContentBounds();
    return { maxCol: bounds.maxCol, maxRow: bounds.maxRow, files, display };
}

/**
 * 序列化元信息（Python `_load` 的第 7 个参数）
 * @param meta - {@link buildScriptMeta} 的结果
 * @returns {string} JSON 字符串
 */
export function buildMetaJson(meta: ScriptMeta): string {
    return JSON.stringify(meta);
}

/**
 * 预加载 `open_grid()` 引用的表格（主线程，异步）
 *
 * 逐张解析路径（相对当前文件所在目录，见 {@link resolveGridKey}）、从 IndexedDB 读存档、
 * 装载成临时 DataCollection，再收集成快照。
 * 解析不到 / 没有存档的路径**静默跳过** —— 该路径不是表格（普通文件 / 打字错误），
 * 此时 Python 侧 `open_grid()` 会带着「可用的表格清单」报错。
 *
 * 返回的 `scratch` 要一直留到运行结束：写回其它表格时直接复用这批实例落存档，
 * 避免「装载两次、解析两次」。
 * @param data - 当前工作表（相对路径的基准目录 + 判定「打开的就是当前表」）
 * @param paths - {@link extractOpenGridPaths} 的结果
 * @returns {Promise<{sheets: OpenedSheet[], scratch: Map<string, DataCollection>}>} 载荷与实例
 */
export async function prepareOpenedSheets(
    data: DataCollection,
    paths: string[],
    options?: { keys?: readonly string[]; preloadAll?: boolean }
): Promise<{ sheets: OpenedSheet[]; scratch: Map<string, DataCollection> }> {
    const sheets: OpenedSheet[] = [];
    const scratch = new Map<string, DataCollection>();
    const byKey = new Map<string, OpenedSheet>();
    const keys = options?.keys ?? await listArchiveKeys(data);

    // 待装载清单：先按用户写的路径解析，再（脚本用到 files 时）把**所有**表格都排进来
    const targets: { input: string; key: string }[] = [];
    if (options?.preloadAll) {
        for (let i = 0; i < keys.length; i++) targets.push({ input: keys[i], key: keys[i] });
    }
    for (let i = 0; i < paths.length; i++) {
        const key = resolveGridKey(keys, data.activedSheetName, paths[i]);
        if (key === null) continue;          // 解析不到 → 跳过（Python 侧 open_grid 会报错）
        targets.push({ input: paths[i], key });
    }

    for (let i = 0; i < targets.length; i++) {
        const { input, key } = targets[i];
        const aliases = buildAliases(input, key);

        // 同一张表的不同写法（"明细.xlsx" 与 "./明细.xlsx"）：只装载一次，别名合并
        const existing = byKey.get(key);
        if (existing !== undefined) {
            for (let a = 0; a < aliases.length; a++) {
                if (existing.aliases.indexOf(aliases[a]) < 0) existing.aliases.push(aliases[a]);
            }
            continue;
        }

        if (key === data.activedSheetName) {
            // 打开的就是当前表：直接让 Python 复用 sheet 对象（写它 = 写当前表）
            const bounds = data.getContentBounds();
            const sheet: OpenedSheet = {
                key, title: key, rowCount: data.rowHeaders.length, colCount: data.colHeaders.length,
                maxCol: bounds.maxCol, maxRow: bounds.maxRow,
                aliases, current: true, text: {}, styles: {}, display: {},
            };
            byKey.set(key, sheet);
            sheets.push(sheet);
            continue;
        }
        const loaded = await DataCollection.loadScratch(key);
        if (loaded === null) continue;       // 该键没有存档（空表 / 读失败）→ 当它不存在
        const snap = collectSnapshot(loaded);
        const bounds = loaded.getContentBounds();
        scratch.set(key, loaded);
        const sheet: OpenedSheet = {
            key, title: key, rowCount: loaded.rowHeaders.length, colCount: loaded.colHeaders.length,
            maxCol: bounds.maxCol, maxRow: bounds.maxRow,
            aliases, current: false, text: snap.text, styles: snap.styles, display: snap.display,
        };
        byKey.set(key, sheet);
        sheets.push(sheet);
    }
    return { sheets, scratch };
}

/**
 * 组装下发给 Python 的「多表」载荷
 * @param sheets - {@link prepareOpenedSheets} 的结果
 * @returns {string} JSON 字符串（Python 侧 `_load` 的第 3 个参数）
 */
export function buildOpenedJson(sheets: OpenedSheet[]): string {
    return JSON.stringify(sheets.map(sheet => (sheet.current
        ? {
            key: sheet.key, title: sheet.title, rows: sheet.rowCount, cols: sheet.colCount,
            maxCol: sheet.maxCol, maxRow: sheet.maxRow, current: true, aliases: sheet.aliases,
        }
        : {
            key: sheet.key, title: sheet.title, rows: sheet.rowCount, cols: sheet.colCount,
            maxCol: sheet.maxCol, maxRow: sheet.maxRow,
            current: false, aliases: sheet.aliases, snapshot: sheet.text, styles: sheet.styles,
            display: sheet.display,
        })));
}
