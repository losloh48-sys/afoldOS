#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyExcel — клон Microsoft Excel на Python
Поддерживает: формулы, диапазоны, форматирование, несколько листов,
графики, сортировку, фильтры, автозаполнение, отмену/повтор.
"""
import tkinter as tk
from tkinter import ttk, colorchooser, messagebox, filedialog, simpledialog
import re
import json
import math
import datetime
import copy
import os
from collections import defaultdict, deque
from typing import Any, Dict, List, Optional, Tuple, Set
# ============================================================================
# КОНСТАНТЫ И УТИЛИТЫ
# ============================================================================
DEFAULT_COL_WIDTH = 100
DEFAULT_ROW_HEIGHT = 22
HEADER_BG = "#f0f0f0"
HEADER_FG = "#000000"
GRID_COLOR = "#d0d0d0"
SELECTION_COLOR = "#3399ff"
SELECTION_FILL = "#cce5ff"
FORMULA_BAR_HEIGHT = 26
NAME_BOX_WIDTH = 80
TAB_HEIGHT = 26
COLORS = {
    "black": "#000000", "white": "#ffffff", "red": "#ff0000",
    "green": "#008000", "blue": "#0000ff", "yellow": "#ffff00",
    "orange": "#ffa500", "purple": "#800080", "gray": "#808080",
    "lightgray": "#c0c0c0", "cyan": "#00ffff", "magenta": "#ff00ff",
}

def col_letter_to_index(letter: str) -> int:
    """Преобразует букву столбца (A, B, ..., Z, AA, AB, ...) в индекс (0-based)."""
    letter = letter.upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1

def col_index_to_letter(index: int) -> str:
    """Преобразует индекс столбца (0-based) в букву."""
    result = ""
    index += 1
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result

def parse_cell_ref(ref: str) -> Optional[Tuple[int, int]]:
    """Парсит ссылку на ячейку (A1, B5, AA100) -> (row, col) 0-based."""
    m = re.match(r'^([A-Za-z]+)(\d+)$', ref.strip())
    if not m:
        return None
    col = col_letter_to_index(m.group(1))
    row = int(m.group(2)) - 1
    if row < 0 or col < 0:
        return None
    return (row, col)

def cell_ref_to_str(row: int, col: int) -> str:
    """Формирует строковое представление ячейки (A1, B5)."""
    return f"{col_index_to_letter(col)}{row + 1}"

def parse_range(range_str: str) -> Optional[List[Tuple[int, int]]]:
    """Парсит диапазон (A1:B5) в список координат."""
    if ':' not in range_str:
        ref = parse_cell_ref(range_str)
        return [ref] if ref else None
    parts = range_str.split(':')
    if len(parts) != 2:
        return None
    start = parse_cell_ref(parts[0])
    end = parse_cell_ref(parts[1])
    if not start or not end:
        return None
    r1, c1 = start
    r2, c2 = end
    r1, r2 = min(r1, r2), max(r1, r2)
    c1, c2 = min(c1, c2), max(c1, c2)
    result = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            result.append((r, c))
    return result

# ============================================================================
# МОДЕЛЬ ДАННЫХ: ЯЧЕЙКА, ЛИСТ, КНИГА
# ============================================================================
class Cell:
    """Модель одной ячейки таблицы."""
    __slots__ = ('value', 'formula', 'cached', 'font_family', 'font_size',
                 'bold', 'italic', 'underline', 'fg_color', 'bg_color',
                 'align', 'border', 'number_format')

    def __init__(self, value: Any = ""):
        self.value: Any = value
        self.formula: Optional[str] = None
        self.cached: Any = None
        self.font_family: str = "Calibri"
        self.font_size: int = 11
        self.bold: bool = False
        self.italic: bool = False
        self.underline: bool = False
        self.fg_color: str = "#000000"
        self.bg_color: Optional[str] = None
        self.align: str = "general"
        self.border: Optional[str] = None
        self.number_format: str = "general"

    def clone(self) -> 'Cell':
        c = Cell(self.value)
        c.formula = self.formula
        c.cached = self.cached
        c.font_family = self.font_family
        c.font_size = self.font_size
        c.bold = self.bold
        c.italic = self.italic
        c.underline = self.underline
        c.fg_color = self.fg_color
        c.bg_color = self.bg_color
        c.align = self.align
        c.border = self.border
        c.number_format = self.number_format
        return c

    def to_dict(self) -> dict:
        d = {}
        if self.value != "":
            d['v'] = self.value
        if self.font_family != "Calibri":
            d['ff'] = self.font_family
        if self.font_size != 11:
            d['fs'] = self.font_size
        if self.bold:
            d['b'] = True
        if self.italic:
            d['i'] = True
        if self.underline:
            d['u'] = True
        if self.fg_color != "#000000":
            d['fg'] = self.fg_color
        if self.bg_color:
            d['bg'] = self.bg_color
        if self.align != "general":
            d['al'] = self.align
        if self.border:
            d['bd'] = self.border
        if self.number_format != "general":
            d['nf'] = self.number_format
        return d

    @classmethod
    def from_dict(cls, d: dict) -> 'Cell':
        c = cls(d.get('v', ''))
        c.font_family = d.get('ff', 'Calibri')
        c.font_size = d.get('fs', 11)
        c.bold = d.get('b', False)
        c.italic = d.get('i', False)
        c.underline = d.get('u', False)
        c.fg_color = d.get('fg', '#000000')
        c.bg_color = d.get('bg')
        c.align = d.get('al', 'general')
        c.border = d.get('bd')
        c.number_format = d.get('nf', 'general')
        return c

class Sheet:
    """Модель одного листа таблицы."""
    MAX_ROWS = 10000
    MAX_COLS = 200

    def __init__(self, name: str = "Sheet1"):
        self.name = name
        self.cells: Dict[Tuple[int, int], Cell] = {}
        self.col_widths: Dict[int, int] = {}
        self.row_heights: Dict[int, int] = {}
        self.frozen_row: int = 0
        self.frozen_col: int = 0
        self.autofilter_range: Optional[Tuple[int, int, int, int]] = None
        self.hidden_rows: Set[int] = set()
        self.hidden_cols: Set[int] = set()

    def get_cell(self, row: int, col: int) -> Cell:
        key = (row, col)
        if key not in self.cells:
            self.cells[key] = Cell()
        return self.cells[key]

    def set_cell_value(self, row: int, col: int, value: Any) -> None:
        cell = self.get_cell(row, col)
        cell.value = value
        cell.formula = None
        cell.cached = None

    def set_cell_formula(self, row: int, col: int, formula: str) -> None:
        cell = self.get_cell(row, col)
        cell.value = formula
        cell.formula = formula
        cell.cached = None

    def get_col_width(self, col: int) -> int:
        return self.col_widths.get(col, DEFAULT_COL_WIDTH)

    def get_row_height(self, row: int) -> int:
        return self.row_heights.get(row, DEFAULT_ROW_HEIGHT)

    def to_dict(self) -> dict:
        cells_dict = {}
        for (r, c), cell in self.cells.items():
            d = cell.to_dict()
            if d:
                cells_dict[f"{cell_ref_to_str(r, c)}"] = d
        return {
            'name': self.name,
            'cells': cells_dict,
            'col_widths': {str(k): v for k, v in self.col_widths.items()},
            'row_heights': {str(k): v for k, v in self.row_heights.items()},
            'frozen_row': self.frozen_row,
            'frozen_col': self.frozen_col,
        }

    @classmethod
    def from_dict(cls, d: dict) -> 'Sheet':
        s = cls(d.get('name', 'Sheet'))
        for ref, cd in d.get('cells', {}).items():
            parsed = parse_cell_ref(ref)
            if parsed:
                r, c = parsed
                s.cells[(r, c)] = Cell.from_dict(cd)
        for k, v in d.get('col_widths', {}).items():
            s.col_widths[int(k)] = v
        for k, v in d.get('row_heights', {}).items():
            s.row_heights[int(k)] = v
        s.frozen_row = d.get('frozen_row', 0)
        s.frozen_col = d.get('frozen_col', 0)
        return s

class Workbook:
    """Модель книги (документа)."""
    def __init__(self):
        self.sheets: List[Sheet] = [Sheet("Sheet1")]
        self.active_index: int = 0
        self.filename: Optional[str] = None
        self.metadata = {
            'created': datetime.datetime.now().isoformat(),
            'modified': datetime.datetime.now().isoformat(),
            'author': 'PyExcel',
        }

    @property
    def active_sheet(self) -> Sheet:
        return self.sheets[self.active_index]

    def add_sheet(self, name: Optional[str] = None) -> Sheet:
        if name is None:
            i = len(self.sheets) + 1
            name = f"Sheet{i}"
            while any(s.name == name for s in self.sheets):
                i += 1
                name = f"Sheet{i}"
        s = Sheet(name)
        self.sheets.append(s)
        return s

    def remove_sheet(self, index: int) -> bool:
        if len(self.sheets) <= 1:
            return False
        if 0 <= index < len(self.sheets):
            self.sheets.pop(index)
            if self.active_index >= len(self.sheets):
                self.active_index = len(self.sheets) - 1
            return True
        return False

    def rename_sheet(self, index: int, new_name: str) -> bool:
        if new_name in [s.name for i, s in enumerate(self.sheets) if i != index]:
            return False
        if 0 <= index < len(self.sheets):
            self.sheets[index].name = new_name
            return True
        return False

    def save(self, filename: str) -> None:
        data = {
            'version': '1.0',
            'metadata': self.metadata,
            'active': self.active_index,
            'sheets': [s.to_dict() for s in self.sheets],
        }
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.filename = filename
        self.metadata['modified'] = datetime.datetime.now().isoformat()

    def load(self, filename: str) -> None:
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.sheets = [Sheet.from_dict(sd) for sd in data.get('sheets', [])]
        if not self.sheets:
            self.sheets = [Sheet("Sheet1")]
        self.active_index = data.get('active', 0)
        self.metadata = data.get('metadata', self.metadata)
        self.filename = filename

# ============================================================================
# ДВИЖОК ФОРМУЛ
# ============================================================================
class FormulaError(Exception):
    pass

class FormulaEngine:
    FUNCTIONS = {
        'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'IF', 'AND', 'OR',
        'NOT', 'CONCAT', 'LEN', 'UPPER', 'LOWER', 'TRIM', 'LEFT', 'RIGHT',
        'MID', 'FIND', 'ABS', 'ROUND', 'CEILING', 'FLOOR', 'SQRT', 'POWER',
        'MOD', 'PI', 'SIN', 'COS', 'TAN', 'LOG', 'LN', 'EXP', 'NOW', 'TODAY',
        'DATE', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'VLOOKUP',
        'HLOOKUP', 'INDEX', 'MATCH', 'COUNTIF', 'SUMIF', 'AVERAGEIF', 'TRUE',
        'FALSE', 'RAND', 'RANDBETWEEN', 'TEXT', 'VALUE', 'INT', 'SIGN',
        'PRODUCT', 'MEDIAN', 'STDEV', 'VAR', 'LARGE', 'SMALL', 'RANK',
        'CONCATENATE', 'REPT', 'SUBSTITUTE', 'PROPER', 'EXACT', 'ISBLANK',
        'ISNUMBER', 'ISTEXT', 'ISERROR', 'IFERROR', 'CHOOSE', 'NA',
    }

    def __init__(self, sheet: Sheet, workbook: Workbook):
        self.sheet = sheet
        self.workbook = workbook
        self.pos = 0
        self.expr = ""
        self.eval_stack: List[Tuple[int, int]] = []
        self.visited: Set[Tuple[int, int]] = set()

    def evaluate(self, formula: str) -> Any:
        if not formula or not formula.startswith('='):
            return formula
        self.expr = formula[1:].strip()
        self.pos = 0
        self.visited = set()
        try:
            result = self.parse_expression()
            self._skip_whitespace()
            if self.pos < len(self.expr):
                raise FormulaError(f"Неожиданный символ: '{self.expr[self.pos]}'")
            return result
        except FormulaError as e:
            return f"#ОШИБКА! {e}"
        except ZeroDivisionError:
            return "#ДЕЛ/0!"
        except Exception:
            return f"#ЗНАЧ!"

    def _skip_whitespace(self):
        while self.pos < len(self.expr) and self.expr[self.pos].isspace():
            self.pos += 1

    def _peek(self) -> Optional[str]:
        self._skip_whitespace()
        if self.pos >= len(self.expr):
            return None
        return self.expr[self.pos]

    def _consume(self, expected: Optional[str] = None) -> str:
        self._skip_whitespace()
        if self.pos >= len(self.expr):
            raise FormulaError("Неожиданный конец формулы")
        ch = self.expr[self.pos]
        if expected is not None and ch != expected:
            raise FormulaError(f"Ожидалось '{expected}', получено '{ch}'")
        self.pos += 1
        return ch

    def parse_expression(self) -> Any:
        return self.parse_comparison()

    def parse_comparison(self) -> Any:
        left = self.parse_additive()
        while True:
            self._skip_whitespace()
            if self.pos >= len(self.expr):
                break
            op = None
            if self.expr[self.pos:self.pos + 2] in ('<=', '>=', '<>'):
                op = self.expr[self.pos:self.pos + 2]
                self.pos += 2
            elif self.expr[self.pos] in ('<', '>', '='):
                op = self.expr[self.pos]
                self.pos += 1
            else:
                break
            right = self.parse_additive()
            left = self._apply_comparison(left, op, right)
        return left

    def _apply_comparison(self, left: Any, op: str, right: Any) -> bool:
        try:
            if isinstance(left, str) and isinstance(right, str):
                l, r = left.lower(), right.lower()
            else:
                l, r = self._to_num(left), self._to_num(right)
        except Exception:
            l, r = str(left), str(right)
        if op == '=': return l == r
        if op == '<>': return l != r
        if op == '<': return l < r
        if op == '>': return l > r
        if op == '<=': return l <= r
        if op == '>=': return l >= r
        return False

    def parse_additive(self) -> Any:
        result = self.parse_multiplicative()
        while True:
            self._skip_whitespace()
            if self.pos >= len(self.expr):
                break
            ch = self.expr[self.pos]
            if ch == '+':
                self.pos += 1
                right = self.parse_multiplicative()
                result = self._to_num(result) + self._to_num(right)
            elif ch == '-':
                self.pos += 1
                right = self.parse_multiplicative()
                result = self._to_num(result) - self._to_num(right)
            elif ch == '&':
                self.pos += 1
                right = self.parse_multiplicative()
                result = str(result if result is not None else "") + str(right if right is not None else "")
            else:
                break
        return result

    def parse_multiplicative(self) -> Any:
        result = self.parse_power()
        while True:
            self._skip_whitespace()
            if self.pos >= len(self.expr):
                break
            ch = self.expr[self.pos]
            if ch == '*':
                self.pos += 1
                right = self.parse_power()
                result = self._to_num(result) * self._to_num(right)
            elif ch == '/':
                self.pos += 1
                right = self.parse_power()
                result = self._to_num(result) / self._to_num(right)
            elif ch == '%':
                self.pos += 1
                result = self._to_num(result) / 100.0
            else:
                break
        return result

    def parse_power(self) -> Any:
        base = self.parse_unary()
        self._skip_whitespace()
        if self.pos < len(self.expr) and self.expr[self.pos] == '^':
            self.pos += 1
            exp = self.parse_unary()
            return self._to_num(base) ** self._to_num(exp)
        return base

    def parse_unary(self) -> Any:
        self._skip_whitespace()
        if self.pos < len(self.expr):
            ch = self.expr[self.pos]
            if ch == '-':
                self.pos += 1
                val = self.parse_primary()
                return -self._to_num(val)
            if ch == '+':
                self.pos += 1
                return self.parse_primary()
        return self.parse_primary()

    def parse_primary(self) -> Any:
        self._skip_whitespace()
        if self.pos >= len(self.expr):
            raise FormulaError("Неожиданный конец формулы")
        ch = self.expr[self.pos]
        if ch == '(':
            self.pos += 1
            result = self.parse_expression()
            self._consume(')')
            return result
        if ch == '"':
            return self._parse_string()
        if ch.isdigit() or ch == '.':
            return self._parse_number()
        if ch.isalpha() or ch == '$' or ch == '_':
            return self._parse_identifier_or_ref()
        raise FormulaError(f"Неожиданный символ: '{ch}'")

    def _parse_string(self) -> str:
        self._consume('"')
        result = []
        while self.pos < len(self.expr):
            ch = self.expr[self.pos]
            if ch == '"':
                if self.pos + 1 < len(self.expr) and self.expr[self.pos + 1] == '"':
                    result.append('"')
                    self.pos += 2
                else:
                    self.pos += 1
                    return ''.join(result)
            else:
                result.append(ch)
                self.pos += 1
        raise FormulaError("Незакрытая строка")

    def _parse_number(self) -> float:
        start = self.pos
        has_dot = False
        while self.pos < len(self.expr):
            ch = self.expr[self.pos]
            if ch.isdigit():
                self.pos += 1
            elif ch == '.' and not has_dot:
                has_dot = True
                self.pos += 1
            elif ch == 'e' or ch == 'E':
                self.pos += 1
                if self.pos < len(self.expr) and self.expr[self.pos] in '+-':
                    self.pos += 1
            else:
                break
        return float(self.expr[start:self.pos])

    def _parse_identifier_or_ref(self) -> Any:
        start = self.pos
        while self.pos < len(self.expr) and (self.expr[self.pos].isalnum() or self.expr[self.pos] in '$_'):
            self.pos += 1
        token = self.expr[start:self.pos]
        self._skip_whitespace()
        if self.pos < len(self.expr) and self.expr[self.pos] == '(' and token.upper() in self.FUNCTIONS:
            return self._call_function(token.upper())
        clean = token.replace('$', '')
        ref = parse_cell_ref(clean)
        if ref:
            return self._eval_cell_ref(ref[0], ref[1])
        if token.upper() == 'TRUE': return True
        if token.upper() == 'FALSE': return False
        if token.upper() == 'PI': return math.pi
        if token.upper() == 'NA': raise FormulaError("#Н/Д")
        raise FormulaError(f"Неизвестный идентификатор: {token}")

    def _eval_cell_ref(self, row: int, col: int) -> Any:
        key = (row, col)
        if key in self.visited:
            raise FormulaError("#ЦИКЛ!")
        self.visited.add(key)
        try:
            target_sheet = self.sheet
            cell = target_sheet.get_cell(row, col)
            if cell.formula:
                sub_engine = FormulaEngine(target_sheet, self.workbook)
                sub_engine.visited = self.visited.copy()
                result = sub_engine.evaluate(cell.formula)
                cell.cached = result
                return result
            if cell.cached is not None:
                return cell.cached
            return self._parse_cell_value(cell.value)
        finally:
            self.visited.discard(key)

    def _parse_cell_value(self, value: Any) -> Any:
        if value == "" or value is None:
            return 0
        if isinstance(value, (int, float, bool)):
            return value
        if isinstance(value, str):
            try:
                if '.' in value:
                    return float(value)
                return int(value)
            except ValueError:
                pass
            if re.match(r'^\d{4}-\d{2}-\d{2}$', value):
                try:
                    return datetime.datetime.strptime(value, '%Y-%m-%d')
                except ValueError:
                    pass
            return value
        return value

    def _to_num(self, value: Any) -> float:
        if value is None or value == "":
            return 0.0
        if isinstance(value, bool):
            return 1.0 if value else 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            try:
                return float(value)
            except ValueError:
                return 0.0
        return 0.0

    def _to_bool(self, value: Any) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return value != 0
        if isinstance(value, str):
            return value.upper() not in ('', '0', 'FALSE', 'ЛОЖЬ')
        return bool(value)

    def _call_function(self, name: str) -> Any:
        self._consume('(')
        args = []
        self._skip_whitespace()
        if self.pos < len(self.expr) and self.expr[self.pos] != ')':
            args.append(self._parse_arg())
            while self.pos < len(self.expr) and self.expr[self.pos] == ',':
                self.pos += 1
                args.append(self._parse_arg())
        self._consume(')')
        return self._execute_function(name, args)

    def _parse_arg(self) -> Any:
        self._skip_whitespace()
        start = self.pos
        ref_start = self.pos
        while self.pos < len(self.expr) and (self.expr[self.pos].isalnum() or self.expr[self.pos] in '$:_'):
            self.pos += 1
        token = self.expr[ref_start:self.pos]
        clean = token.replace('$', '')
        self._skip_whitespace()
        if self.pos < len(self.expr) and self.expr[self.pos] == ':':
            self.pos += 1
            ref_end_start = self.pos
            while self.pos < len(self.expr) and (self.expr[self.pos].isalnum() or self.expr[self.pos] in '$'):
                self.pos += 1
            end_token = self.expr[ref_end_start:self.pos]
            end_clean = end_token.replace('$', '')
            start_ref = parse_cell_ref(clean)
            end_ref = parse_cell_ref(end_clean)
            if start_ref and end_ref:
                return self._eval_range(start_ref, end_ref)
        self.pos = start
        return self.parse_expression()

    def _eval_range(self, start: Tuple[int, int], end: Tuple[int, int]) -> List[Any]:
        r1, c1 = start
        r2, c2 = end
        r1, r2 = min(r1, r2), max(r1, r2)
        c1, c2 = min(c1, c2), max(c1, c2)
        result = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                val = self._eval_cell_ref(r, c)
                result.append(val)
        return result

    def _flatten_args(self, args: List[Any]) -> List[Any]:
        result = []
        for arg in args:
            if isinstance(arg, list):
                result.extend(arg)
            else:
                result.append(arg)
        return result

    def _numeric_args(self, args: List[Any]) -> List[float]:
        flat = self._flatten_args(args)
        result = []
        for v in flat:
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                result.append(float(v))
            elif isinstance(v, str):
                try:
                    result.append(float(v))
                except ValueError:
                    pass
        return result

    def _execute_function(self, name: str, args: List[Any]) -> Any:
        if name == 'SUM':
            return sum(self._numeric_args(args))
        if name == 'AVERAGE':
            nums = self._numeric_args(args)
            if not nums:
                raise FormulaError("#ДЕЛ/0!")
            return sum(nums) / len(nums)
        if name == 'MIN':
            nums = self._numeric_args(args)
            if not nums:
                return 0
            return min(nums)
        if name == 'MAX':
            nums = self._numeric_args(args)
            if not nums:
                return 0
            return max(nums)
        if name == 'COUNT':
            flat = self._flatten_args(args)
            return sum(1 for v in flat if isinstance(v, (int, float)) and not isinstance(v, bool))
        if name == 'COUNTA':
            flat = self._flatten_args(args)
            return sum(1 for v in flat if v is not None and v != "")
        if name == 'PRODUCT':
            nums = self._numeric_args(args)
            if not nums:
                return 0
            result = 1
            for n in nums:
                result *= n
            return result
        if name == 'MEDIAN':
            nums = sorted(self._numeric_args(args))
            if not nums:
                return 0
            n = len(nums)
            if n % 2 == 1:
                return nums[n // 2]
            return (nums[n // 2 - 1] + nums[n // 2]) / 2
        if name == 'STDEV':
            nums = self._numeric_args(args)
            if len(nums) < 2:
                raise FormulaError("#ДЕЛ/0!")
            mean = sum(nums) / len(nums)
            variance = sum((x - mean) ** 2 for x in nums) / (len(nums) - 1)
            return math.sqrt(variance)
        if name == 'VAR':
            nums = self._numeric_args(args)
            if len(nums) < 2:
                raise FormulaError("#ДЕЛ/0!")
            mean = sum(nums) / len(nums)
            return sum((x - mean) ** 2 for x in nums) / (len(nums) - 1)
        if name == 'LARGE':
            nums = sorted(self._numeric_args(args[:-1]), reverse=True)
            k = int(self._to_num(args[-1]))
            if k < 1 or k > len(nums):
                raise FormulaError("#ЧИСЛО!")
            return nums[k - 1]
        if name == 'SMALL':
            nums = sorted(self._numeric_args(args[:-1]))
            k = int(self._to_num(args[-1]))
            if k < 1 or k > len(nums):
                raise FormulaError("#ЧИСЛО!")
            return nums[k - 1]
        if name == 'RANK':
            num = self._to_num(args[0])
            nums = self._numeric_args([args[1]])
            order = int(self._to_num(args[2])) if len(args) > 2 else 0
            sorted_nums = sorted(nums, reverse=(order == 0))
            try:
                return sorted_nums.index(num) + 1
            except ValueError:
                raise FormulaError("#Н/Д")
        if name == 'ABS':
            return abs(self._to_num(args[0]))
        if name == 'ROUND':
            num = self._to_num(args[0])
            digits = int(self._to_num(args[1])) if len(args) > 1 else 0
            return round(num, digits)
        if name == 'CEILING':
            num = self._to_num(args[0])
            sig = self._to_num(args[1]) if len(args) > 1 else 1
            if sig == 0:
                return 0
            return math.ceil(num / sig) * sig
        if name == 'FLOOR':
            num = self._to_num(args[0])
            sig = self._to_num(args[1]) if len(args) > 1 else 1
            if sig == 0:
                return 0
            return math.floor(num / sig) * sig
        if name == 'INT':
            return math.floor(self._to_num(args[0]))
        if name == 'SIGN':
            n = self._to_num(args[0])
            return 1 if n > 0 else (-1 if n < 0 else 0)
        if name == 'SQRT':
            return math.sqrt(self._to_num(args[0]))
        if name == 'POWER':
            return self._to_num(args[0]) ** self._to_num(args[1])
        if name == 'MOD':
            return self._to_num(args[0]) % self._to_num(args[1])
        if name == 'PI':
            return math.pi
        if name == 'SIN':
            return math.sin(self._to_num(args[0]))
        if name == 'COS':
            return math.cos(self._to_num(args[0]))
        if name == 'TAN':
            return math.tan(self._to_num(args[0]))
        if name == 'LOG':
            num = self._to_num(args[0])
            base = self._to_num(args[1]) if len(args) > 1 else 10
            return math.log(num, base)
        if name == 'LN':
            return math.log(self._to_num(args[0]))
        if name == 'EXP':
            return math.exp(self._to_num(args[0]))
        if name == 'RAND':
            import random
            return random.random()
        if name == 'RANDBETWEEN':
            import random
            lo = int(self._to_num(args[0]))
            hi = int(self._to_num(args[1]))
            return random.randint(lo, hi)
        if name == 'IF':
            if len(args) < 2:
                raise FormulaError("IF требует минимум 2 аргумента")
            cond = self._to_bool(args[0])
            if cond:
                return args[1]
            return args[2] if len(args) > 2 else False
        if name == 'IFERROR':
            try:
                return args[0]
            except Exception:
                return args[1] if len(args) > 1 else ""
        if name == 'AND':
            flat = self._flatten_args(args)
            return all(self._to_bool(v) for v in flat)
        if name == 'OR':
            flat = self._flatten_args(args)
            return any(self._to_bool(v) for v in flat)
        if name == 'NOT':
            return not self._to_bool(args[0])
        if name == 'TRUE':
            return True
        if name == 'FALSE':
            return False
        if name == 'CONCAT' or name == 'CONCATENATE':
            flat = self._flatten_args(args)
            return ''.join(str(v) if v is not None else '' for v in flat)
        if name == 'LEN':
            return len(str(args[0]))
        if name == 'UPPER':
            return str(args[0]).upper()
        if name == 'LOWER':
            return str(args[0]).lower()
        if name == 'PROPER':
            return str(args[0]).title()
        if name == 'TRIM':
            return str(args[0]).strip()
        if name == 'LEFT':
            s = str(args[0])
            n = int(self._to_num(args[1])) if len(args) > 1 else 1
            return s[:n]
        if name == 'RIGHT':
            s = str(args[0])
            n = int(self._to_num(args[1])) if len(args) > 1 else 1
            return s[-n:] if n > 0 else ""
        if name == 'MID':
            s = str(args[0])
            start = int(self._to_num(args[1])) - 1
            n = int(self._to_num(args[2]))
            return s[start:start + n]
        if name == 'FIND':
            needle = str(args[0])
            haystack = str(args[1])
            start = int(self._to_num(args[2])) - 1 if len(args) > 2 else 0
            idx = haystack.find(needle, start)
            if idx == -1:
                raise FormulaError("#ЗНАЧ!")
            return idx + 1
        if name == 'REPT':
            return str(args[0]) * int(self._to_num(args[1]))
        if name == 'SUBSTITUTE':
            s = str(args[0])
            old = str(args[1])
            new = str(args[2])
            return s.replace(old, new)
        if name == 'EXACT':
            return str(args[0]) == str(args[1])
        if name == 'ISBLANK':
            return args[0] is None or args[0] == ""
        if name == 'ISNUMBER':
            return isinstance(args[0], (int, float)) and not isinstance(args[0], bool)
        if name == 'ISTEXT':
            return isinstance(args[0], str)
        if name == 'ISERROR':
            return isinstance(args[0], str) and args[0].startswith('#')
        if name == 'VALUE':
            return self._to_num(args[0])
        if name == 'TEXT':
            val = args[0]
            fmt = str(args[1]) if len(args) > 1 else "general"
            if fmt == "general":
                return str(val)
            if isinstance(val, (int, float)):
                try:
                    if '.' in fmt:
                        decimals = len(fmt.split('.')[-1])
                        return f"{val:.{decimals}f}"
                    return str(int(val))
                except Exception:
                    return str(val)
            return str(val)
        if name == 'NOW':
            return datetime.datetime.now()
        if name == 'TODAY':
            return datetime.datetime.now().date()
        if name == 'DATE':
            y = int(self._to_num(args[0]))
            m = int(self._to_num(args[1]))
            d = int(self._to_num(args[2]))
            return datetime.date(y, m, d)
        if name == 'YEAR':
            v = args[0]
            if isinstance(v, (datetime.date, datetime.datetime)):
                return v.year
            raise FormulaError("#ЗНАЧ!")
        if name == 'MONTH':
            v = args[0]
            if isinstance(v, (datetime.date, datetime.datetime)):
                return v.month
            raise FormulaError("#ЗНАЧ!")
        if name == 'DAY':
            v = args[0]
            if isinstance(v, (datetime.date, datetime.datetime)):
                return v.day
            raise FormulaError("#ЗНАЧ!")
        if name == 'HOUR':
            v = args[0]
            if isinstance(v, datetime.datetime):
                return v.hour
            raise FormulaError("#ЗНАЧ!")
        if name == 'MINUTE':
            v = args[0]
            if isinstance(v, datetime.datetime):
                return v.minute
            raise FormulaError("#ЗНАЧ!")
        if name == 'SECOND':
            v = args[0]
            if isinstance(v, datetime.datetime):
                return v.second
            raise FormulaError("#ЗНАЧ!")
        if name == 'VLOOKUP':
            key = args[0]
            table = args[1] if isinstance(args[1], list) else self._flatten_args([args[1]])
            col_idx = int(self._to_num(args[2])) - 1
            if not isinstance(table, list):
                raise FormulaError("#Н/Д")
            for i, v in enumerate(table):
                if v == key or (isinstance(v, (int, float)) and isinstance(key, (int, float)) and abs(v - key) < 1e-9):
                    if i + col_idx < len(table):
                        return table[i + col_idx]
            raise FormulaError("#Н/Д")
        if name == 'HLOOKUP':
            key = args[0]
            table = args[1] if isinstance(args[1], list) else self._flatten_args([args[1]])
            row_idx = int(self._to_num(args[2])) - 1
            for i, v in enumerate(table):
                if v == key:
                    if i + row_idx < len(table):
                        return table[i + row_idx]
            raise FormulaError("#Н/Д")
        if name == 'INDEX':
            arr = args[0] if isinstance(args[0], list) else self._flatten_args([args[0]])
            row = int(self._to_num(args[1])) - 1 if len(args) > 1 else 0
            if 0 <= row < len(arr):
                return arr[row]
            raise FormulaError("#ССЫЛКА!")
        if name == 'MATCH':
            key = args[0]
            arr = args[1] if isinstance(args[1], list) else self._flatten_args([args[1]])
            for i, v in enumerate(arr):
                if v == key:
                    return i + 1
            raise FormulaError("#Н/Д")
        if name == 'COUNTIF':
            rng = args[0] if isinstance(args[0], list) else self._flatten_args([args[0]])
            criteria = args[1]
            return sum(1 for v in rng if self._matches_criteria(v, criteria))
        if name == 'SUMIF':
            rng = args[0] if isinstance(args[0], list) else self._flatten_args([args[0]])
            criteria = args[1]
            sum_rng = args[2] if len(args) > 2 else rng
            if not isinstance(sum_rng, list):
                sum_rng = self._flatten_args([sum_rng])
            total = 0.0
            for i, v in enumerate(rng):
                if self._matches_criteria(v, criteria):
                    if i < len(sum_rng):
                        total += self._to_num(sum_rng[i])
            return total
        if name == 'AVERAGEIF':
            rng = args[0] if isinstance(args[0], list) else self._flatten_args([args[0]])
            criteria = args[1]
            sum_rng = args[2] if len(args) > 2 else rng
            if not isinstance(sum_rng, list):
                sum_rng = self._flatten_args([sum_rng])
            total = 0.0
            count = 0
            for i, v in enumerate(rng):
                if self._matches_criteria(v, criteria):
                    if i < len(sum_rng):
                        total += self._to_num(sum_rng[i])
                    count += 1
            if count == 0:
                raise FormulaError("#ДЕЛ/0!")
            return total / count
        if name == 'CHOOSE':
            idx = int(self._to_num(args[0]))
            if 1 <= idx < len(args):
                return args[idx]
            raise FormulaError("#ЗНАЧ!")
        raise FormulaError(f"Неизвестная функция: {name}")

    def _matches_criteria(self, value: Any, criteria: Any) -> bool:
        if isinstance(criteria, str):
            m = re.match(r'^([<>=!]+)(.+)$', criteria)
            if m:
                op, val_str = m.groups()
                try:
                    val = float(val_str)
                except ValueError:
                    val = val_str
                if op in ('=', '=='): return value == val
                if op in ('<>', '!='): return value != val
                if op == '<': return self._to_num(value) < self._to_num(val)
                if op == '>': return self._to_num(value) > self._to_num(val)
                if op == '<=': return self._to_num(value) <= self._to_num(val)
                if op == '>=': return self._to_num(value) >= self._to_num(val)
            return str(value).lower() == criteria.lower()
        return value == criteria

# ============================================================================
# МЕНЕДЖЕР ИСТОРИИ (UNDO/REDO)
# ============================================================================
class HistoryManager:
    MAX_HISTORY = 100

    def __init__(self):
        self.undo_stack: deque = deque(maxlen=self.MAX_HISTORY)
        self.redo_stack: deque = deque(maxlen=self.MAX_HISTORY)

    def push(self, action: dict) -> None:
        self.undo_stack.append(action)
        self.redo_stack.clear()

    def can_undo(self) -> bool:
        return bool(self.undo_stack)

    def can_redo(self) -> bool:
        return bool(self.redo_stack)

    def pop_undo(self) -> Optional[dict]:
        if self.undo_stack:
            action = self.undo_stack.pop()
            self.redo_stack.append(action)
            return action
        return None

    def pop_redo(self) -> Optional[dict]:
        if self.redo_stack:
            action = self.redo_stack.pop()
            self.undo_stack.append(action)
            return action
        return None

    def clear(self) -> None:
        self.undo_stack.clear()
        self.redo_stack.clear()

# ============================================================================
# ГЛАВНОЕ ПРИЛОЖЕНИЕ
# ============================================================================
class SpreadsheetApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PyExcel — Python Spreadsheet")
        self.root.geometry("1280x800")
        self.root.minsize(800, 600)

        self.workbook = Workbook()
        self.history = HistoryManager()

        self.selection_start: Optional[Tuple[int, int]] = None
        self.selection_end: Optional[Tuple[int, int]] = None
        self.active_cell: Tuple[int, int] = (0, 0)
        self.editing: bool = False
        self.edit_buffer: str = ""
        self.cursor_pos: int = 0

        self.scroll_row: int = 0
        self.scroll_col: int = 0

        self.dragging: bool = False
        self.drag_start: Optional[Tuple[int, int]] = None
        self.resizing_col: Optional[int] = None
        self.resizing_row: Optional[int] = None
        self.resize_start_pos: int = 0
        self.resize_start_size: int = 0

        self.clipboard: Optional[Dict[Tuple[int, int], Cell]] = None
        self.clipboard_range: Optional[Tuple[int, int, int, int]] = None

        self.autofill_active: bool = False
        self.autofill_start: Optional[Tuple[int, int]] = None
        self.autofill_end: Optional[Tuple[int, int]] = None

        self.blink_visible: bool = True
        self.cursor_blink_id: Optional[str] = None

        self.show_grid: bool = True
        self.zoom: float = 1.0

        self._build_ui()
        self._bind_events()
        self._recalculate_all()
        self._redraw()
        self._start_cursor_blink()

    def _build_ui(self):
        self._build_menu()
        self._build_toolbar()
        self._build_formula_bar()

        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(self.main_frame, bg="white", highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.v_scroll = tk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self._on_vscroll)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scroll = tk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self._on_hscroll)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self._build_sheet_tabs()

        self.status_bar = tk.Label(self.root, text="Готово", anchor=tk.W,
                                   bg="#f0f0f0", padx=8, pady=2)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Новый", command=self._new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="Открыть...", command=self._open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Сохранить", command=self._save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Сохранить как...", command=self._save_as)
        file_menu.add_separator()
        file_menu.add_command(label="Экспорт в CSV...", command=self._export_csv)
        file_menu.add_command(label="Импорт CSV...", command=self._import_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit)
        menubar.add_cascade(label="Файл", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Отменить", command=self._undo, accelerator="Ctrl+Z")
        edit_menu.add_command(label="Повторить", command=self._redo, accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="Вырезать", command=self._cut, accelerator="Ctrl+X")
        edit_menu.add_command(label="Копировать", command=self._copy, accelerator="Ctrl+C")
        edit_menu.add_command(label="Вставить", command=self._paste, accelerator="Ctrl+V")
        edit_menu.add_command(label="Очистить", command=self._clear_selection, accelerator="Del")
        edit_menu.add_separator()
        edit_menu.add_command(label="Выделить всё", command=self._select_all, accelerator="Ctrl+A")
        edit_menu.add_command(label="Найти...", command=self._find, accelerator="Ctrl+F")
        edit_menu.add_command(label="Заменить...", command=self._replace, accelerator="Ctrl+H")
        menubar.add_cascade(label="Правка", menu=edit_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        self._show_grid_var = tk.BooleanVar(value=True)
        view_menu.add_checkbutton(label="Показывать сетку", command=self._toggle_grid,
                                  variable=self._show_grid_var)
        view_menu.add_separator()
        view_menu.add_command(label="Увеличить", command=lambda: self._set_zoom(self.zoom + 0.1))
        view_menu.add_command(label="Уменьшить", command=lambda: self._set_zoom(self.zoom - 0.1))
        view_menu.add_command(label="100%", command=lambda: self._set_zoom(1.0))
        menubar.add_cascade(label="Вид", menu=view_menu)

        insert_menu = tk.Menu(menubar, tearoff=0)
        insert_menu.add_command(label="Функция...", command=self._insert_function)
        insert_menu.add_command(label="Дата и время", command=self._insert_datetime)
        insert_menu.add_command(label="Диаграмма...", command=self._insert_chart)
        menubar.add_cascade(label="Вставка", menu=insert_menu)

        format_menu = tk.Menu(menubar, tearoff=0)
        format_menu.add_command(label="Шрифт...", command=self._format_font)
        format_menu.add_command(label="Цвет ячейки...", command=self._format_fill)
        format_menu.add_command(label="Цвет текста...", command=self._format_text_color)
        format_menu.add_command(label="Границы...", command=self._format_border)
        format_menu.add_separator()
        format_menu.add_command(label="Числовой формат...", command=self._format_number)
        format_menu.add_command(label="Объединить ячейки", command=self._merge_cells)
        menubar.add_cascade(label="Формат", menu=format_menu)

        data_menu = tk.Menu(menubar, tearoff=0)
        data_menu.add_command(label="Сортировка по возрастанию", command=lambda: self._sort_column(True))
        data_menu.add_command(label="Сортировка по убыванию", command=lambda: self._sort_column(False))
        data_menu.add_separator()
        data_menu.add_command(label="Фильтр", command=self._toggle_filter)
        data_menu.add_command(label="Удалить фильтр", command=self._clear_filter)
        data_menu.add_separator()
        data_menu.add_command(label="Текст по столбцам...", command=self._text_to_columns)
        data_menu.add_command(label="Удалить дубликаты", command=self._remove_duplicates)
        menubar.add_cascade(label="Данные", menu=data_menu)

        sheet_menu = tk.Menu(menubar, tearoff=0)
        sheet_menu.add_command(label="Новый лист", command=self._add_sheet)
        sheet_menu.add_command(label="Переименовать...", command=self._rename_sheet)
        sheet_menu.add_command(label="Удалить лист", command=self._delete_sheet)
        menubar.add_cascade(label="Лист", menu=sheet_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="О программе", command=self._show_about)
        help_menu.add_command(label="Поддерживаемые функции", command=self._show_functions_help)
        menubar.add_cascade(label="Справка", menu=help_menu)

    def _build_toolbar(self):
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        toolbar.pack(fill=tk.X, padx=2, pady=2)

        self.font_var = tk.StringVar(value="Calibri")
        fonts = ["Arial", "Calibri", "Times New Roman", "Courier New", "Verdana",
                 "Georgia", "Tahoma", "Comic Sans MS"]
        font_cb = ttk.Combobox(toolbar, textvariable=self.font_var, values=fonts, width=14)
        font_cb.pack(side=tk.LEFT, padx=2)
        font_cb.bind("<<ComboboxSelected>>", self._on_font_change)

        self.size_var = tk.StringVar(value="11")
        size_cb = ttk.Combobox(toolbar, textvariable=self.size_var,
                               values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"],
                               width=4)
        size_cb.pack(side=tk.LEFT, padx=2)
        size_cb.bind("<<ComboboxSelected>>", self._on_size_change)

        # ИСПРАВЛЕНО: ttk.Separator вместо tk.Separator
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        self.bold_var = tk.BooleanVar()
        self.italic_var = tk.BooleanVar()
        self.underline_var = tk.BooleanVar()
        tk.Checkbutton(toolbar, text="Ж", variable=self.bold_var,
                       command=self._on_bold_change, font=("Calibri", 10, "bold"),
                       width=2).pack(side=tk.LEFT, padx=1)
        tk.Checkbutton(toolbar, text="К", variable=self.italic_var,
                       command=self._on_italic_change, font=("Calibri", 10, "italic"),
                       width=2).pack(side=tk.LEFT, padx=1)
        tk.Checkbutton(toolbar, text="Ч", variable=self.underline_var,
                       command=self._on_underline_change, font=("Calibri", 10, "underline"),
                       width=2).pack(side=tk.LEFT, padx=1)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        self.align_var = tk.StringVar(value="general")
        tk.Radiobutton(toolbar, text="⬅", variable=self.align_var, value="left",
                       command=self._on_align_change, indicatoron=False, width=2).pack(side=tk.LEFT, padx=1)
        tk.Radiobutton(toolbar, text="⬌", variable=self.align_var, value="center",
                       command=self._on_align_change, indicatoron=False, width=2).pack(side=tk.LEFT, padx=1)
        tk.Radiobutton(toolbar, text="➡", variable=self.align_var, value="right",
                       command=self._on_align_change, indicatoron=False, width=2).pack(side=tk.LEFT, padx=1)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        # ИСПРАВЛЕНО: bg вместо toolbackground
        tk.Button(toolbar, text="🎨", command=self._format_fill, width=2,
                  bg="#f0f0f0").pack(side=tk.LEFT, padx=1)
        tk.Button(toolbar, text="A", command=self._format_text_color, width=2,
                  fg="red", font=("Calibri", 10, "bold")).pack(side=tk.LEFT, padx=1)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        self.num_format_var = tk.StringVar(value="general")
        fmt_cb = ttk.Combobox(toolbar, textvariable=self.num_format_var,
                              values=["general", "number", "currency", "percent", "date", "text"],
                              width=10)
        fmt_cb.pack(side=tk.LEFT, padx=2)
        fmt_cb.bind("<<ComboboxSelected>>", self._on_num_format_change)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4)

        tk.Button(toolbar, text="Σ", command=self._auto_sum, width=2,
                  font=("Calibri", 10, "bold")).pack(side=tk.LEFT, padx=1)
        tk.Button(toolbar, text="📊", command=self._insert_chart, width=2).pack(side=tk.LEFT, padx=1)
        tk.Button(toolbar, text="🔽", command=lambda: self._sort_column(True), width=2).pack(side=tk.LEFT, padx=1)
        tk.Button(toolbar, text="🔼", command=lambda: self._sort_column(False), width=2).pack(side=tk.LEFT, padx=1)

    def _build_formula_bar(self):
        bar = tk.Frame(self.root, bd=1, relief=tk.SUNKEN)
        bar.pack(fill=tk.X, padx=2, pady=1)

        self.name_var = tk.StringVar(value="A1")
        self.name_box = tk.Entry(bar, textvariable=self.name_var, width=8,
                                 font=("Calibri", 10), justify=tk.CENTER)
        self.name_box.pack(side=tk.LEFT, padx=2)
        self.name_box.bind("<Return>", self._on_name_enter)

        tk.Label(bar, text="fx", font=("Calibri", 11, "italic"), fg="#666").pack(side=tk.LEFT, padx=4)

        self.formula_var = tk.StringVar()
        self.formula_entry = tk.Entry(bar, textvariable=self.formula_var,
                                      font=("Calibri", 10))
        self.formula_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.formula_entry.bind("<Return>", self._on_formula_enter)
        self.formula_entry.bind("<Escape>", self._on_formula_escape)
        self.formula_entry.bind("<KeyRelease>", self._on_formula_key)

    def _build_sheet_tabs(self):
        self.tab_frame = tk.Frame(self.root, height=TAB_HEIGHT, bg="#e0e0e0")
        self.tab_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.tab_frame.pack_propagate(False)

        tk.Button(self.tab_frame, text="+", command=self._add_sheet,
                  width=2, bg="#f0f0f0").pack(side=tk.LEFT, padx=2, pady=2)

        self.tab_container = tk.Frame(self.tab_frame, bg="#e0e0e0")
        self.tab_container.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._refresh_tabs()

    def _bind_events(self):
        self.canvas.bind("<Button-1>", self._on_mouse_down)
        self.canvas.bind("<B1-Motion>", self._on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_mouse_up)
        self.canvas.bind("<Double-Button-1>", self._on_double_click)
        self.canvas.bind("<Button-3>", self._on_right_click)
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.canvas.bind("<Configure>", self._on_resize)

        # The spreadsheet receives keyboard input through its Canvas. Bind there
        # instead of EmbeddedRoot so shortcuts work without leaking into AFold.
        self.canvas.focus_set()
        self.canvas.bind("<Key>", self._on_key)
        self.canvas.bind("<Return>", self._on_enter)
        self.canvas.bind("<Escape>", self._on_escape)
        self.canvas.bind("<BackSpace>", self._on_backspace)
        self.canvas.bind("<Delete>", self._on_delete)
        self.canvas.bind("<Tab>", self._on_tab)
        self.canvas.bind("<Up>", lambda e:self._move_selection(0,-1))
        self.canvas.bind("<Down>", lambda e:self._move_selection(0,1))
        self.canvas.bind("<Left>", lambda e:self._move_selection(-1,0))
        self.canvas.bind("<Right>", lambda e:self._move_selection(1,0))
        self.canvas.bind("<Home>", lambda e:self._move_to_start())
        self.canvas.bind("<End>", lambda e:self._move_to_end())
        self.canvas.bind("<Control-a>", lambda e:(self._select_all(),"break")[1])
        self.canvas.bind("<Control-c>", lambda e:(self._copy(),"break")[1])
        self.canvas.bind("<Control-x>", lambda e:(self._cut(),"break")[1])
        self.canvas.bind("<Control-v>", lambda e:(self._paste(),"break")[1])
        self.canvas.bind("<Control-z>", lambda e:(self._undo(),"break")[1])
        self.canvas.bind("<Control-y>", lambda e:(self._redo(),"break")[1])
        self.canvas.bind("<Control-s>", lambda e:(self._save_file(),"break")[1])
        self.canvas.bind("<Control-o>", lambda e:(self._open_file(),"break")[1])
        self.canvas.bind("<Control-n>", lambda e:(self._new_file(),"break")[1])
        self.canvas.bind("<Control-f>", lambda e:(self._find(),"break")[1])
        self.canvas.bind("<Control-h>", lambda e:(self._replace(),"break")[1])
        self.canvas.bind("<Control-b>", lambda e:(self._toggle_bold(),"break")[1])
        self.canvas.bind("<Control-i>", lambda e:(self._toggle_italic(),"break")[1])
        self.canvas.bind("<Control-u>", lambda e:(self._toggle_underline(),"break")[1])

    def _redraw(self):
        self.canvas.delete("all")
        sheet = self.workbook.active_sheet
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        if width <= 1 or height <= 1:
            return
        zoom = self.zoom
        header_h = int(24 * zoom)
        row_header_w = int(50 * zoom)

        x = row_header_w
        col = self.scroll_col
        visible_cols = []
        while x < width and col < sheet.MAX_COLS:
            w = int(sheet.get_col_width(col) * zoom)
            if col in sheet.hidden_cols:
                col += 1
                continue
            self.canvas.create_rectangle(x, 0, x + w, header_h,
                                         fill=HEADER_BG, outline=GRID_COLOR)
            label = col_index_to_letter(col)
            self.canvas.create_text(x + w // 2, header_h // 2, text=label,
                                    font=("Calibri", 9), fill=HEADER_FG)
            visible_cols.append((col, x, w))
            x += w
            col += 1

        y = header_h
        row = self.scroll_row
        visible_rows = []
        while y < height and row < sheet.MAX_ROWS:
            h = int(sheet.get_row_height(row) * zoom)
            if row in sheet.hidden_rows:
                row += 1
                continue
            self.canvas.create_rectangle(0, y, row_header_w, y + h,
                                         fill=HEADER_BG, outline=GRID_COLOR)
            self.canvas.create_text(row_header_w // 2, y + h // 2, text=str(row + 1),
                                    font=("Calibri", 9), fill=HEADER_FG)
            visible_rows.append((row, y, h))
            y += h
            row += 1

        self.canvas.create_rectangle(0, 0, row_header_w, header_h,
                                     fill=HEADER_BG, outline=GRID_COLOR)

        for row, y, rh in visible_rows:
            for col, x, cw in visible_cols:
                self._draw_cell(row, col, x, y, cw, rh)

        self._draw_selection(visible_rows, visible_cols)
        if self.editing:
            self._draw_edit_cursor(visible_rows, visible_cols)

        self._visible_rows = visible_rows
        self._visible_cols = visible_cols
        self._header_h = header_h
        self._row_header_w = row_header_w

        self._update_scrollbars()

    def _draw_cell(self, row: int, col: int, x: int, y: int, w: int, h: int):
        sheet = self.workbook.active_sheet
        cell = sheet.cells.get((row, col))

        bg = cell.bg_color if cell and cell.bg_color else "white"
        self.canvas.create_rectangle(x, y, x + w, y + h, fill=bg, outline="")

        if self.show_grid:
            self.canvas.create_line(x + w, y, x + w, y + h, fill=GRID_COLOR)
            self.canvas.create_line(x, y + h, x + w, y + h, fill=GRID_COLOR)

        if cell and cell.border:
            border_w = {'thin': 1, 'medium': 2, 'thick': 3}.get(cell.border, 1)
            self.canvas.create_rectangle(x, y, x + w, y + h, outline="black", width=border_w)

        if cell is None:
            return
        display_val = self._get_display_value(cell)
        if not display_val:
            return

        align = cell.align
        if align == "general":
            if isinstance(cell.cached, (int, float)):
                align = "right"
            else:
                align = "left"
        anchor = {'left': 'w', 'center': 'center', 'right': 'e'}.get(align, 'w')
        text_x = {'left': x + 4, 'center': x + w // 2, 'right': x + w - 4}.get(align, x + 4)

        style = ""
        if cell.bold: style += "bold "
        if cell.italic: style += "italic "
        if cell.underline: style += "underline "
        font_size = int(cell.font_size * self.zoom)
        font = (cell.font_family, font_size, style.strip())

        max_width = w - 8
        text = display_val
        try:
            tmp_id = self.canvas.create_text(0, 0, text=text, font=font, anchor='w')
            text_width = self.canvas.bbox(tmp_id)
            self.canvas.delete(tmp_id)
            if text_width and (text_width[2] - text_width[0]) > max_width:
                while len(text) > 1:
                    text = text[:-1]
                    tmp_id = self.canvas.create_text(0, 0, text=text + "…", font=font, anchor='w')
                    text_width = self.canvas.bbox(tmp_id)
                    self.canvas.delete(tmp_id)
                    if text_width and (text_width[2] - text_width[0]) <= max_width:
                        text = text + "…"
                        break
        except Exception:
            pass

        self.canvas.create_text(text_x, y + h // 2, text=text, font=font,
                                fill=cell.fg_color, anchor=anchor)

    def _get_display_value(self, cell: Cell) -> str:
        if cell.formula:
            val = cell.cached
        else:
            val = cell.value
        if val is None or val == "":
            return ""
        if isinstance(val, bool):
            return "ИСТИНА" if val else "ЛОЖЬ"
        if isinstance(val, datetime.datetime):
            return val.strftime("%d.%m.%Y %H:%M:%S")
        if isinstance(val, datetime.date):
            return val.strftime("%d.%m.%Y")
        if isinstance(val, float):
            if cell.number_format == "percent":
                return f"{val * 100:.2f}%"
            if cell.number_format == "currency":
                return f"{val:,.2f} ₽"
            if cell.number_format == "number":
                return f"{val:,.2f}"
            if val == int(val):
                return str(int(val))
            return f"{val:g}"
        if isinstance(val, int):
            return str(val)
        return str(val)

    def _draw_selection(self, visible_rows, visible_cols):
        if self.selection_start is None or self.selection_end is None:
            return
        r1, c1 = self.selection_start
        r2, c2 = self.selection_end
        r1, r2 = min(r1, r2), max(r1, r2)
        c1, c2 = min(c1, c2), max(c1, c2)

        x1 = y1 = x2 = y2 = None
        for col, x, w in visible_cols:
            if c1 <= col <= c2:
                if x1 is None: x1 = x
                x2 = x + w
        for row, y, h in visible_rows:
            if r1 <= row <= r2:
                if y1 is None: y1 = y
                y2 = y + h
        if x1 is None or y1 is None:
            return

        self.canvas.create_rectangle(x1, y1, x2, y2, fill=SELECTION_FILL, outline="")
        self.canvas.create_rectangle(x1, y1, x2, y2, outline=SELECTION_COLOR, width=2)

        for col, x, w in visible_cols:
            if col == self.active_cell[1]:
                ax1, ax2 = x, x + w
                break
        else:
            return
        for row, y, h in visible_rows:
            if row == self.active_cell[0]:
                ay1, ay2 = y, y + h
                break
        else:
            return
        self.canvas.create_rectangle(ax1, ay1, ax2, ay2, outline=SELECTION_COLOR, width=2)

        marker_size = 6
        self.canvas.create_rectangle(x2 - marker_size, y2 - marker_size, x2, y2,
                                     fill=SELECTION_COLOR, outline="black")

    def _draw_edit_cursor(self, visible_rows, visible_cols):
        if not self.blink_visible:
            return
        row, col = self.active_cell
        for c, x, w in visible_cols:
            if c == col:
                cx1, cw = x, w
                break
        else:
            return
        for r, y, h in visible_rows:
            if r == row:
                cy1, ch = y, h
                break
        else:
            return

        cell = self.workbook.active_sheet.get_cell(row, col)
        text_before = self.edit_buffer[:self.cursor_pos]
        font_size = int(cell.font_size * self.zoom)
        font = (cell.font_family, font_size)
        try:
            tmp_id = self.canvas.create_text(0, 0, text=text_before, font=font, anchor='w')
            text_width = self.canvas.bbox(tmp_id)
            self.canvas.delete(tmp_id)
            cursor_x = cx1 + 4 + (text_width[2] - text_width[0] if text_width else 0)
        except Exception:
            cursor_x = cx1 + 4
        cursor_y1 = cy1 + 3
        cursor_y2 = cy1 + ch - 3
        self.canvas.create_line(cursor_x, cursor_y1, cursor_x, cursor_y2,
                                fill="black", width=1)

    def _start_cursor_blink(self):
        def blink():
            if self.editing:
                self.blink_visible = not self.blink_visible
                self._redraw()
            self.cursor_blink_id = self.root.after(500, blink)
        self.cursor_blink_id = self.root.after(500, blink)

    def _hit_test(self, x: int, y: int) -> Tuple[str, Any]:
        if not hasattr(self, '_visible_rows') or not hasattr(self, '_visible_cols'):
            return ("none", None)
        header_h = self._header_h
        row_header_w = self._row_header_w

        if x < row_header_w and y < header_h:
            return ("corner", None)

        if y < header_h:
            for col, cx, cw in self._visible_cols:
                if cx <= x < cx + cw:
                    if abs(x - (cx + cw)) < 3:
                        return ("col_resize", col)
                    return ("col_header", col)
            return ("none", None)

        if x < row_header_w:
            for row, ry, rh in self._visible_rows:
                if ry <= y < ry + rh:
                    if abs(y - (ry + rh)) < 3:
                        return ("row_resize", row)
                    return ("row_header", row)
            return ("none", None)

        for row, ry, rh in self._visible_rows:
            if ry <= y < ry + rh:
                for col, cx, cw in self._visible_cols:
                    if cx <= x < cx + cw:
                        if (self.selection_start and self.selection_end and
                                abs(x - (cx + cw)) < 6 and abs(y - (ry + rh)) < 6):
                            return ("autofill_handle", (row, col))
                        return ("cell", (row, col))
        return ("none", None)

    def _on_mouse_down(self, event):
        target, data = self._hit_test(event.x, event.y)
        if target == "cell":
            row, col = data
            if event.state & 0x1:
                self.selection_end = (row, col)
            else:
                self.active_cell = (row, col)
                self.selection_start = (row, col)
                self.selection_end = (row, col)
            self._commit_edit()
            self._update_name_box()
            self._update_formula_bar()
            self._redraw()
        elif target == "col_header":
            col = data
            self.active_cell = (self.active_cell[0], col)
            self.selection_start = (0, col)
            self.selection_end = (Sheet.MAX_ROWS - 1, col)
            self._commit_edit()
            self._redraw()
        elif target == "row_header":
            row = data
            self.active_cell = (row, self.active_cell[1])
            self.selection_start = (row, 0)
            self.selection_end = (row, Sheet.MAX_COLS - 1)
            self._commit_edit()
            self._redraw()
        elif target == "col_resize":
            self.resizing_col = data
            self.resize_start_pos = event.x
            sheet = self.workbook.active_sheet
            self.resize_start_size = sheet.get_col_width(data)
        elif target == "row_resize":
            self.resizing_row = data
            self.resize_start_pos = event.y
            sheet = self.workbook.active_sheet
            self.resize_start_size = sheet.get_row_height(data)
        elif target == "autofill_handle":
            self.autofill_active = True
            self.autofill_start = self.selection_start
            self.autofill_end = self.selection_end
        elif target == "corner":
            self._select_all()

    def _on_mouse_drag(self, event):
        if self.resizing_col is not None:
            delta = event.x - self.resize_start_pos
            new_size = max(20, self.resize_start_size + int(delta / self.zoom))
            self.workbook.active_sheet.col_widths[self.resizing_col] = new_size
            self._redraw()
            return
        if self.resizing_row is not None:
            delta = event.y - self.resize_start_pos
            new_size = max(15, self.resize_start_size + int(delta / self.zoom))
            self.workbook.active_sheet.row_heights[self.resizing_row] = new_size
            self._redraw()
            return
        if self.autofill_active:
            target, data = self._hit_test(event.x, event.y)
            if target == "cell":
                self.autofill_end = data
                self._redraw()
            return
        target, data = self._hit_test(event.x, event.y)
        if target == "cell":
            self.selection_end = data
            self._redraw()

    def _on_mouse_up(self, event):
        if self.resizing_col is not None:
            self.resizing_col = None
            return
        if self.resizing_row is not None:
            self.resizing_row = None
            return
        if self.autofill_active:
            self.autofill_active = False
            self._perform_autofill()
            return

    def _on_double_click(self, event):
        target, data = self._hit_test(event.x, event.y)
        if target == "cell":
            row, col = data
            self.active_cell = (row, col)
            self.selection_start = (row, col)
            self.selection_end = (row, col)
            self._start_edit()
        elif target == "col_resize":
            self._auto_fit_column(data)

    def _on_right_click(self, event):
        target, data = self._hit_test(event.x, event.y)
        menu = tk.Menu(self.root, tearoff=0)
        if target == "cell":
            menu.add_command(label="Вырезать", command=self._cut, accelerator="Ctrl+X")
            menu.add_command(label="Копировать", command=self._copy, accelerator="Ctrl+C")
            menu.add_command(label="Вставить", command=self._paste, accelerator="Ctrl+V")
            menu.add_separator()
            menu.add_command(label="Вставить функцию...", command=self._insert_function)
            menu.add_separator()
            menu.add_command(label="Формат ячейки...", command=self._format_font)
            menu.add_separator()
            menu.add_command(label="Очистить содержимое", command=self._clear_selection)
        elif target == "col_header":
            menu.add_command(label="Вставить столбец слева", command=lambda: self._insert_col(data))
            menu.add_command(label="Вставить столбец справа", command=lambda: self._insert_col(data + 1))
            menu.add_command(label="Удалить столбец", command=lambda: self._delete_col(data))
            menu.add_separator()
            menu.add_command(label="Автоподбор ширины", command=lambda: self._auto_fit_column(data))
            menu.add_command(label="Ширина столбца...", command=lambda: self._set_col_width(data))
        elif target == "row_header":
            menu.add_command(label="Вставить строку выше", command=lambda: self._insert_row(data))
            menu.add_command(label="Вставить строку ниже", command=lambda: self._insert_row(data + 1))
            menu.add_command(label="Удалить строку", command=lambda: self._delete_row(data))
            menu.add_separator()
            menu.add_command(label="Высота строки...", command=lambda: self._set_row_height(data))

        if len(menu.children) > 0 or menu.index("end") is not None:
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

    def _on_mouse_wheel(self, event):
        delta = -1 if event.delta > 0 else 1
        self.scroll_row = max(0, self.scroll_row + delta * 3)
        self._redraw()

    def _on_resize(self, event):
        self._redraw()

    def _on_key(self, event):
        if event.state & 0x4:
            return
        if event.keysym in ('Return', 'Escape', 'BackSpace', 'Delete', 'Tab',
                            'Up', 'Down', 'Left', 'Right', 'Home', 'End',
                            'Shift_L', 'Shift_R', 'Control_L', 'Control_R',
                            'Alt_L', 'Alt_R', 'Caps_Lock'):
            return
        if not event.char:
            return
        if not self.editing:
            self._start_edit(event.char)
        else:
            self.edit_buffer = (self.edit_buffer[:self.cursor_pos] +
                                event.char +
                                self.edit_buffer[self.cursor_pos:])
            self.cursor_pos += 1
            self.formula_var.set('=' + self.edit_buffer if self.edit_buffer.startswith('=') else self.edit_buffer)
            self._redraw()

    def _on_enter(self, event):
        if self.formula_entry.focus_get() == self.formula_entry:
            self._on_formula_enter(event)
            return
        if self.name_box.focus_get() == self.name_box:
            self._on_name_enter(event)
            return
        if self.editing:
            self._commit_edit()
            self._move_selection(0, 1)
        else:
            self._start_edit()

    def _on_escape(self, event):
        if self.editing:
            self._cancel_edit()

    def _on_backspace(self, event):
        if self.editing:
            if self.cursor_pos > 0:
                self.edit_buffer = (self.edit_buffer[:self.cursor_pos - 1] +
                                    self.edit_buffer[self.cursor_pos:])
                self.cursor_pos -= 1
                self.formula_var.set(self.edit_buffer)
                self._redraw()
        else:
            self._clear_selection()

    def _on_delete(self, event):
        if not self.editing:
            self._clear_selection()

    def _on_tab(self, event):
        if self.editing:
            self._commit_edit()
        if event.state & 0x1:
            self._move_selection(-1, 0)
        else:
            self._move_selection(1, 0)
        return "break"

    def _start_edit(self, initial: str = ""):
        row, col = self.active_cell
        cell = self.workbook.active_sheet.get_cell(row, col)
        if initial:
            self.edit_buffer = initial
            self.editing = True
            self.cursor_pos = len(initial)
        elif cell.value != "" or cell.formula:
            self.edit_buffer = str(cell.value)
            self.editing = True
            self.cursor_pos = len(self.edit_buffer)
        else:
            self.edit_buffer = ""
            self.editing = True
            self.cursor_pos = 0
        self.formula_var.set(self.edit_buffer)
        self.formula_entry.focus_set()
        self.blink_visible = True
        self._redraw()

    def _commit_edit(self):
        if not self.editing:
            return
        row, col = self.active_cell
        sheet = self.workbook.active_sheet
        old_value = sheet.get_cell(row, col).value
        new_value = self.edit_buffer

        if old_value != new_value:
            self.history.push({
                'type': 'edit',
                'sheet': self.workbook.active_index,
                'row': row, 'col': col,
                'old': old_value,
                'new': new_value,
            })

        if new_value.startswith('='):
            sheet.set_cell_formula(row, col, new_value)
        else:
            sheet.set_cell_value(row, col, new_value)

        self.editing = False
        self.edit_buffer = ""
        self.cursor_pos = 0
        self._recalculate_all()
        self._redraw()
        self.root.focus_set()

    def _cancel_edit(self):
        self.editing = False
        self.edit_buffer = ""
        self.cursor_pos = 0
        self._update_formula_bar()
        self._redraw()
        self.root.focus_set()

    def _on_formula_enter(self, event):
        if self.editing:
            self._commit_edit()
            self._move_selection(0, 1)
            return "break"

    def _on_formula_escape(self, event):
        self._cancel_edit()
        return "break"

    def _on_formula_key(self, event):
        if self.editing:
            self.edit_buffer = self.formula_var.get()
            self.cursor_pos = len(self.edit_buffer)

    def _on_name_enter(self, event):
        ref = self.name_var.get().strip()
        parsed = parse_cell_ref(ref)
        if parsed:
            row, col = parsed
            self.active_cell = (row, col)
            self.selection_start = (row, col)
            self.selection_end = (row, col)
            self._scroll_to_cell(row, col)
            self._update_formula_bar()
            self._redraw()
            return "break"

    def _move_selection(self, dx: int, dy: int):
        if self.editing:
            self._commit_edit()
        row, col = self.active_cell
        new_row = max(0, min(Sheet.MAX_ROWS - 1, row + dy))
        new_col = max(0, min(Sheet.MAX_COLS - 1, col + dx))
        self.active_cell = (new_row, new_col)
        self.selection_start = (new_row, new_col)
        self.selection_end = (new_row, new_col)
        self._scroll_to_cell(new_row, new_col)
        self._update_name_box()
        self._update_formula_bar()
        self._redraw()

    def _move_to_start(self):
        if self.editing:
            self.cursor_pos = 0
            self._redraw()
        else:
            self._move_selection(-self.active_cell[1], 0)

    def _move_to_end(self):
        if self.editing:
            self.cursor_pos = len(self.edit_buffer)
            self._redraw()
        else:
            self._move_selection(Sheet.MAX_COLS - self.active_cell[1] - 1, 0)

    def _scroll_to_cell(self, row: int, col: int):
        sheet = self.workbook.active_sheet
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        header_h = int(24 * self.zoom)
        row_header_w = int(50 * self.zoom)

        x = 0
        for c in range(self.scroll_col, col + 1):
            if c == col:
                if x + sheet.get_col_width(c) * self.zoom > width - row_header_w:
                    self.scroll_col = col
                    break
                x += sheet.get_col_width(c) * self.zoom
        else:
            if col < self.scroll_col:
                self.scroll_col = col

        y = 0
        for r in range(self.scroll_row, row + 1):
            if r == row:
                if y + sheet.get_row_height(r) * self.zoom > height - header_h:
                    self.scroll_row = row
                    break
                y += sheet.get_row_height(r) * self.zoom
        else:
            if row < self.scroll_row:
                self.scroll_row = row

    def _select_all(self):
        self.selection_start = (0, 0)
        self.selection_end = (Sheet.MAX_ROWS - 1, Sheet.MAX_COLS - 1)
        self._redraw()

    def _get_selection_bounds(self) -> Tuple[int, int, int, int]:
        if not self.selection_start or not self.selection_end:
            r, c = self.active_cell
            return (r, c, r, c)
        r1, c1 = self.selection_start
        r2, c2 = self.selection_end
        return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))

    def _update_name_box(self):
        row, col = self.active_cell
        self.name_var.set(cell_ref_to_str(row, col))

    def _update_formula_bar(self):
        row, col = self.active_cell
        cell = self.workbook.active_sheet.get_cell(row, col)
        if cell.value != "" or cell.formula:
            self.formula_var.set(str(cell.value))
        else:
            self.formula_var.set("")

        self.font_var.set(cell.font_family)
        self.size_var.set(str(cell.font_size))
        self.bold_var.set(cell.bold)
        self.italic_var.set(cell.italic)
        self.underline_var.set(cell.underline)
        self.align_var.set(cell.align)
        self.num_format_var.set(cell.number_format)

    def _update_status(self, text: str):
        self.status_bar.config(text=text)

    def _update_scrollbars(self):
        sheet = self.workbook.active_sheet
        total_rows = max(100, max((r for r, c in sheet.cells.keys()), default=0) + 50)
        total_cols = max(26, max((c for r, c in sheet.cells.keys()), default=0) + 20)
        if total_rows > 0:
            self.v_scroll.set(self.scroll_row / total_rows,
                              (self.scroll_row + 30) / total_rows)
        if total_cols > 0:
            self.h_scroll.set(self.scroll_col / total_cols,
                              (self.scroll_col + 15) / total_cols)

    def _on_vscroll(self, *args):
        command, value = args[0], args[1]
        if command == "moveto":
            self.scroll_row = int(float(value) * 1000)
        elif command == "scroll":
            self.scroll_row = max(0, self.scroll_row + int(value) * 3)
        self._redraw()

    def _on_hscroll(self, *args):
        command, value = args[0], args[1]
        if command == "moveto":
            self.scroll_col = int(float(value) * 100)
        elif command == "scroll":
            self.scroll_col = max(0, self.scroll_col + int(value))
        self._redraw()

    def _recalculate_all(self):
        sheet = self.workbook.active_sheet
        engine = FormulaEngine(sheet, self.workbook)
        for (r, c), cell in list(sheet.cells.items()):
            if cell.formula:
                cell.cached = engine.evaluate(cell.formula)

    def _new_file(self):
        if messagebox.askyesno("Новый файл", "Создать новый документ? Несохранённые изменения будут потеряны."):
            self.workbook = Workbook()
            self.history.clear()
            self.scroll_row = 0
            self.scroll_col = 0
            self.active_cell = (0, 0)
            self.selection_start = (0, 0)
            self.selection_end = (0, 0)
            self._refresh_tabs()
            self._update_name_box()
            self._update_formula_bar()
            self._redraw()

    def _open_file(self):
        filename = filedialog.askopenfilename(
            title="Открыть файл",
            filetypes=[("PyExcel files", "*.pyxl"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                self.workbook.load(filename)
                self.history.clear()
                self.scroll_row = 0
                self.scroll_col = 0
                self.active_cell = (0, 0)
                self.selection_start = (0, 0)
                self.selection_end = (0, 0)
                self._recalculate_all()
                self._refresh_tabs()
                self._update_name_box()
                self._update_formula_bar()
                self._redraw()
                self._update_status(f"Открыт: {filename}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

    def load_file_from_path(self, filename):
        """Open a workbook supplied by AFold Explorer/file associations."""
        try:
            self.workbook.load(str(filename))
            self.history.clear()
            self.scroll_row = 0
            self.scroll_col = 0
            self.active_cell = (0, 0)
            self.selection_start = (0, 0)
            self.selection_end = (0, 0)
            self._recalculate_all()
            self._refresh_tabs()
            self._update_name_box()
            self._update_formula_bar()
            self._redraw()
            self._update_status(f"Открыт: {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка открытия", str(e), parent=self.root)

    def _save_file(self):
        if self.workbook.filename:
            try:
                self.workbook.save(self.workbook.filename)
                self._update_status(f"Сохранено: {self.workbook.filename}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")
        else:
            self._save_as()

    def _save_as(self):
        filename = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".pyxl",
            filetypes=[("PyExcel files", "*.pyxl"), ("JSON files", "*.json")]
        )
        if filename:
            try:
                self.workbook.save(filename)
                self._update_status(f"Сохранено: {filename}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")

    def _export_csv(self):
        filename = filedialog.asksaveasfilename(
            title="Экспорт в CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not filename:
            return
        try:
            import csv
            sheet = self.workbook.active_sheet
            max_row = max((r for r, c in sheet.cells.keys()), default=0)
            max_col = max((c for r, c in sheet.cells.keys()), default=0)
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                for r in range(max_row + 1):
                    row_data = []
                    for c in range(max_col + 1):
                        cell = sheet.cells.get((r, c))
                        if cell:
                            val = cell.cached if cell.formula else cell.value
                            row_data.append(val if val != "" else "")
                        else:
                            row_data.append("")
                    writer.writerow(row_data)
            self._update_status(f"Экспортировано: {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать:\n{e}")

    def _import_csv(self):
        filename = filedialog.askopenfilename(
            title="Импорт CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not filename:
            return
        try:
            import csv
            sheet = self.workbook.add_sheet(os.path.basename(filename))
            self.workbook.active_index = len(self.workbook.sheets) - 1
            with open(filename, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for r, row_data in enumerate(reader):
                    for c, val in enumerate(row_data):
                        if val:
                            sheet.set_cell_value(r, c, val)
            self._recalculate_all()
            self._refresh_tabs()
            self._redraw()
            self._update_status(f"Импортировано: {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать:\n{e}")

    def _undo(self):
        action = self.history.pop_undo()
        if not action:
            return
        t = action['type']
        if t == 'edit':
            sheet = self.workbook.sheets[action['sheet']]
            cell = sheet.get_cell(action['row'], action['col'])
            cell.value = action['old']
            cell.formula = None
            cell.cached = None
            self._recalculate_all()
            self._redraw()
        elif t == 'paste':
            sheet = self.workbook.sheets[action['sheet']]
            for (r, c), old_val in action['old_values'].items():
                cell = sheet.get_cell(r, c)
                cell.value = old_val
                cell.formula = None
                cell.cached = None
            self._recalculate_all()
            self._redraw()
        elif t == 'format':
            sheet = self.workbook.sheets[action['sheet']]
            for (r, c), old_fmt in action['old_formats'].items():
                cell = sheet.get_cell(r, c)
                for k, v in old_fmt.items():
                    setattr(cell, k, v)
            self._redraw()

    def _redo(self):
        action = self.history.pop_redo()
        if not action:
            return
        t = action['type']
        if t == 'edit':
            sheet = self.workbook.sheets[action['sheet']]
            cell = sheet.get_cell(action['row'], action['col'])
            if str(action['new']).startswith('='):
                cell.value = action['new']
                cell.formula = action['new']
            else:
                cell.value = action['new']
                cell.formula = None
            cell.cached = None
            self._recalculate_all()
            self._redraw()

    def _cut(self):
        self._copy()
        self._clear_selection()

    def _copy(self):
        r1, c1, r2, c2 = self._get_selection_bounds()
        self.clipboard = {}
        sheet = self.workbook.active_sheet
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = sheet.cells.get((r, c))
                if cell and (cell.value != "" or cell.formula):
                    self.clipboard[(r - r1, c - c1)] = cell.clone()
        self.clipboard_range = (r1, c1, r2, c2)
        self._update_status(f"Скопировано: {r2 - r1 + 1}×{c2 - c1 + 1}")

    def _paste(self):
        if not self.clipboard:
            return
        row, col = self.active_cell
        sheet = self.workbook.active_sheet
        old_values = {}
        for (dr, dc), cell in self.clipboard.items():
            r, c = row + dr, col + dc
            target = sheet.get_cell(r, c)
            old_values[(r, c)] = target.value
            new_cell = cell.clone()
            if new_cell.formula:
                r1, c1 = self.clipboard_range[:2]
                dr_orig = row - r1
                dc_orig = col - c1
                new_cell.formula = self._shift_formula(new_cell.formula, dr_orig, dc_orig)
                new_cell.value = new_cell.formula
                new_cell.cached = None
            sheet.cells[(r, c)] = new_cell
        self.history.push({
            'type': 'paste',
            'sheet': self.workbook.active_index,
            'old_values': old_values,
        })
        self._recalculate_all()
        self._redraw()

    def _shift_formula(self, formula: str, dr: int, dc: int) -> str:
        if dr == 0 and dc == 0:
            return formula

        def replacer(m):
            ref = m.group(0)
            parts = re.findall(r'\$?[A-Za-z]+\$?\d+', ref)
            if not parts:
                return ref
            result = ref
            for part in parts:
                clean = part.replace('$', '')
                parsed = parse_cell_ref(clean)
                if not parsed:
                    continue
                r, c = parsed
                dollar_positions = [i for i, ch in enumerate(part) if ch == '$']
                col_abs = 0 in dollar_positions
                row_abs = (len(part) - 1) in dollar_positions or (len(part) > 1 and part[-2] == '$')
                new_r = r if row_abs else r + dr
                new_c = c if col_abs else c + dc
                if new_r < 0 or new_c < 0:
                    continue
                letter = col_index_to_letter(new_c)
                num = new_r + 1
                new_ref = ('' if not col_abs else '$') + letter + ('' if not row_abs else '$') + str(num)
                result = result.replace(part, new_ref, 1)
            return result

        return re.sub(r'\$?[A-Za-z]+\$?\d+', replacer, formula)

    def _clear_selection(self):
        r1, c1, r2, c2 = self._get_selection_bounds()
        sheet = self.workbook.active_sheet
        old_values = {}
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = sheet.cells.get((r, c))
                if cell and cell.value != "":
                    old_values[(r, c)] = cell.value
                    cell.value = ""
                    cell.formula = None
                    cell.cached = None
        if old_values:
            self.history.push({
                'type': 'paste',
                'sheet': self.workbook.active_index,
                'old_values': old_values,
            })
        self._recalculate_all()
        self._redraw()

    def _apply_format(self, attr: str, value: Any):
        r1, c1, r2, c2 = self._get_selection_bounds()
        sheet = self.workbook.active_sheet
        old_formats = {}
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = sheet.get_cell(r, c)
                old_formats[(r, c)] = {attr: getattr(cell, attr)}
                setattr(cell, attr, value)
        self.history.push({
            'type': 'format',
            'sheet': self.workbook.active_index,
            'old_formats': old_formats,
        })
        self._redraw()

    def _on_font_change(self, event=None):
        self._apply_format('font_family', self.font_var.get())

    def _on_size_change(self, event=None):
        try:
            self._apply_format('font_size', int(self.size_var.get()))
        except ValueError:
            pass

    def _on_bold_change(self):
        self._apply_format('bold', self.bold_var.get())

    def _on_italic_change(self):
        self._apply_format('italic', self.italic_var.get())

    def _on_underline_change(self):
        self._apply_format('underline', self.underline_var.get())

    def _on_align_change(self):
        self._apply_format('align', self.align_var.get())

    def _on_num_format_change(self, event=None):
        self._apply_format('number_format', self.num_format_var.get())

    def _toggle_bold(self):
        self.bold_var.set(not self.bold_var.get())
        self._on_bold_change()

    def _toggle_italic(self):
        self.italic_var.set(not self.italic_var.get())
        self._on_italic_change()

    def _toggle_underline(self):
        self.underline_var.set(not self.underline_var.get())
        self._on_underline_change()

    def _format_font(self):
        from tkinter import font as tkfont
        row, col = self.active_cell
        cell = self.workbook.active_sheet.get_cell(row, col)
        dialog = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
        dialog.title("Формат шрифта")
        dialog.geometry("300x250")
        dialog.transient(self.root)
        dialog.grab_set()
        tk.Label(dialog, text="Шрифт:").pack(anchor=tk.W, padx=10, pady=5)
        font_var = tk.StringVar(value=cell.font_family)
        ttk.Combobox(dialog, textvariable=font_var,
                     values=["Arial", "Calibri", "Times New Roman", "Courier New", "Verdana"]).pack(fill=tk.X, padx=10)
        tk.Label(dialog, text="Размер:").pack(anchor=tk.W, padx=10, pady=5)
        size_var = tk.StringVar(value=str(cell.font_size))
        ttk.Combobox(dialog, textvariable=size_var,
                     values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24"]).pack(fill=tk.X, padx=10)
        bold_var = tk.BooleanVar(value=cell.bold)
        tk.Checkbutton(dialog, text="Жирный", variable=bold_var).pack(anchor=tk.W, padx=10, pady=5)
        italic_var = tk.BooleanVar(value=cell.italic)
        tk.Checkbutton(dialog, text="Курсив", variable=italic_var).pack(anchor=tk.W, padx=10)

        def apply():
            self._apply_format('font_family', font_var.get())
            try:
                self._apply_format('font_size', int(size_var.get()))
            except ValueError:
                pass
            self._apply_format('bold', bold_var.get())
            self._apply_format('italic', italic_var.get())
            dialog.destroy()

        tk.Button(dialog, text="OK", command=apply).pack(pady=10)

    def _format_fill(self):
        color = colorchooser.askcolor(title="Цвет фона")[1]
        if color:
            self._apply_format('bg_color', color)

    def _format_text_color(self):
        color = colorchooser.askcolor(title="Цвет текста")[1]
        if color:
            self._apply_format('fg_color', color)

    def _format_border(self):
        dialog = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
        dialog.title("Границы")
        dialog.geometry("200x150")
        dialog.transient(self.root)
        tk.Label(dialog, text="Тип границы:").pack(pady=5)
        border_var = tk.StringVar(value="none")
        for val, text in [("none", "Нет"), ("thin", "Тонкая"), ("medium", "Средняя"), ("thick", "Толстая")]:
            tk.Radiobutton(dialog, text=text, variable=border_var, value=val).pack(anchor=tk.W, padx=20)

        def apply():
            val = border_var.get()
            self._apply_format('border', None if val == "none" else val)
            dialog.destroy()

        tk.Button(dialog, text="OK", command=apply).pack(pady=10)

    def _format_number(self):
        dialog = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
        dialog.title("Числовой формат")
        dialog.geometry("250x200")
        dialog.transient(self.root)
        tk.Label(dialog, text="Формат:").pack(pady=5)
        fmt_var = tk.StringVar(value="general")
        for val, text in [("general", "Общий"), ("number", "Число (2 знака)"),
                          ("currency", "Денежный"), ("percent", "Процентный"),
                          ("date", "Дата"), ("text", "Текст")]:
            tk.Radiobutton(dialog, text=text, variable=fmt_var, value=val).pack(anchor=tk.W, padx=20)

        def apply():
            self._apply_format('number_format', fmt_var.get())
            dialog.destroy()

        tk.Button(dialog, text="OK", command=apply).pack(pady=10)

    def _merge_cells(self):
        messagebox.showinfo("Информация",
                            "Объединение ячеек: в данной версии отображается только верхняя левая ячейка выделения.")

    def _sort_column(self, ascending: bool = True):
        row, col = self.active_cell
        sheet = self.workbook.active_sheet
        max_row = 0
        for r in range(row, Sheet.MAX_ROWS):
            if any((r, c) in sheet.cells and sheet.cells[(r, c)].value != ""
                   for c in range(Sheet.MAX_COLS)):
                max_row = r
            else:
                break
        if max_row <= row:
            return

        rows_data = []
        for r in range(row, max_row + 1):
            row_cells = {}
            for c in range(Sheet.MAX_COLS):
                cell = sheet.cells.get((r, c))
                if cell and cell.value != "":
                    row_cells[c] = cell.clone()
            rows_data.append((r, row_cells))

        def sort_key(item):
            cells = item[1]
            if col in cells:
                val = cells[col].cached if cells[col].formula else cells[col].value
                if isinstance(val, (int, float)):
                    return (0, val, "")
                return (1, 0, str(val).lower())
            return (2, 0, "")

        rows_data.sort(key=sort_key, reverse=not ascending)

        old_values = {}
        for new_r, (_, cells) in enumerate(rows_data, start=row):
            for c in range(Sheet.MAX_COLS):
                old_cell = sheet.cells.get((new_r, c))
                if old_cell and old_cell.value != "":
                    old_values[(new_r, c)] = old_cell.value
                if c in cells:
                    sheet.cells[(new_r, c)] = cells[c]
                elif (new_r, c) in sheet.cells:
                    del sheet.cells[(new_r, c)]

        self._recalculate_all()
        self._redraw()
        self._update_status(f"Отсортировано по столбцу {col_index_to_letter(col)}")

    def _toggle_filter(self):
        sheet = self.workbook.active_sheet
        if sheet.autofilter_range:
            self._clear_filter()
            return
        r1, c1, r2, c2 = self._get_selection_bounds()
        sheet.autofilter_range = (r1, c1, r2, c2)
        self._update_status("Фильтр включён")
        self._redraw()

    def _clear_filter(self):
        sheet = self.workbook.active_sheet
        sheet.autofilter_range = None
        sheet.hidden_rows.clear()
        self._redraw()

    def _text_to_columns(self):
        messagebox.showinfo("Информация", "Текст по столбцам: разделитель — запятая/табуляция.")

    def _remove_duplicates(self):
        r1, c1, r2, c2 = self._get_selection_bounds()
        sheet = self.workbook.active_sheet
        seen = set()
        to_delete = []
        for r in range(r1, r2 + 1):
            key = []
            for c in range(c1, c2 + 1):
                cell = sheet.cells.get((r, c))
                val = cell.value if cell else ""
                key.append(str(val))
            key = tuple(key)
            if key in seen:
                to_delete.append(r)
            else:
                seen.add(key)
        if to_delete:
            for r in reversed(to_delete):
                for c in range(c1, c2 + 1):
                    if (r, c) in sheet.cells:
                        del sheet.cells[(r, c)]
        self._update_status(f"Удалено дубликатов: {len(to_delete)}")
        self._redraw()

    def _refresh_tabs(self):
        for w in self.tab_container.winfo_children():
            w.destroy()
        for i, sheet in enumerate(self.workbook.sheets):
            btn = tk.Button(self.tab_container, text=sheet.name,
                            command=lambda idx=i: self._switch_sheet(idx),
                            bg="white" if i == self.workbook.active_index else "#e0e0e0",
                            padx=10, pady=2)
            btn.pack(side=tk.LEFT, padx=1, pady=2)

    def _switch_sheet(self, index: int):
        self.workbook.active_index = index
        self.active_cell = (0, 0)
        self.selection_start = (0, 0)
        self.selection_end = (0, 0)
        self.scroll_row = 0
        self.scroll_col = 0
        self._refresh_tabs()
        self._update_name_box()
        self._update_formula_bar()
        self._redraw()

    def _add_sheet(self):
        name = simpledialog.askstring("Новый лист", "Имя листа:",
                                      initialvalue=f"Sheet{len(self.workbook.sheets) + 1}")
        if name:
            if any(s.name == name for s in self.workbook.sheets):
                messagebox.showerror("Ошибка", "Лист с таким именем уже существует")
                return
            self.workbook.add_sheet(name)
            self.workbook.active_index = len(self.workbook.sheets) - 1
            self._refresh_tabs()
            self._redraw()

    def _rename_sheet(self):
        name = simpledialog.askstring("Переименовать лист", "Новое имя:",
                                      initialvalue=self.workbook.active_sheet.name)
        if name:
            if not self.workbook.rename_sheet(self.workbook.active_index, name):
                messagebox.showerror("Ошибка", "Не удалось переименовать")
            else:
                self._refresh_tabs()

    def _delete_sheet(self):
        if len(self.workbook.sheets) <= 1:
            messagebox.showwarning("Предупреждение", "Нельзя удалить единственный лист")
            return
        if messagebox.askyesno("Удалить лист",
                               f"Удалить лист '{self.workbook.active_sheet.name}'?"):
            self.workbook.remove_sheet(self.workbook.active_index)
            self._refresh_tabs()
            self._redraw()

    def _insert_function(self):
        dialog = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
        dialog.title("Вставка функции")
        dialog.geometry("400x500")
        dialog.transient(self.root)
        tk.Label(dialog, text="Выберите функцию:").pack(pady=5)
        listbox = tk.Listbox(dialog, height=15)
        functions = sorted(FormulaEngine.FUNCTIONS)
        for f in functions:
            listbox.insert(tk.END, f)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tk.Label(dialog, text="Описание:").pack(anchor=tk.W, padx=10)
        desc_text = tk.Text(dialog, height=5, wrap=tk.WORD)
        desc_text.pack(fill=tk.X, padx=10, pady=5)
        descriptions = {
            'SUM': 'SUM(число1, [число2], ...) — сумма аргументов',
            'AVERAGE': 'AVERAGE(число1, ...) — среднее значение',
            'IF': 'IF(условие; значение_если_истина; значение_если_ложь)',
            'VLOOKUP': 'VLOOKUP(искомое; таблица; номер_столбца; [приближение])',
            'COUNT': 'COUNT(значение1, ...) — подсчёт чисел',
            'MAX': 'MAX(число1, ...) — максимальное значение',
            'MIN': 'MIN(число1, ...) — минимальное значение',
        }

        def on_select(event):
            sel = listbox.curselection()
            if sel:
                name = listbox.get(sel[0])
                desc_text.delete(1.0, tk.END)
                desc_text.insert(1.0, descriptions.get(name, f"Функция {name}"))

        listbox.bind("<<ListboxSelect>>", on_select)

        def insert():
            sel = listbox.curselection()
            if sel:
                name = listbox.get(sel[0])
                if not self.editing:
                    self._start_edit(f"={name}(")
                else:
                    self.edit_buffer += f"{name}("
                    self.cursor_pos = len(self.edit_buffer)
                    self.formula_var.set(self.edit_buffer)
                dialog.destroy()

        tk.Button(dialog, text="Вставить", command=insert).pack(pady=5)

    def _insert_datetime(self):
        now = datetime.datetime.now()
        if not self.editing:
            self._start_edit(now.strftime("%Y-%m-%d %H:%M:%S"))
        else:
            self.edit_buffer += now.strftime("%Y-%m-%d %H:%M:%S")
            self.cursor_pos = len(self.edit_buffer)
            self.formula_var.set(self.edit_buffer)
            self._redraw()

    def _insert_chart(self):
        r1, c1, r2, c2 = self._get_selection_bounds()
        sheet = self.workbook.active_sheet
        labels = []
        values = []
        if c2 - c1 == 1:
            for r in range(r1, r2 + 1):
                lbl_cell = sheet.cells.get((r, c1))
                val_cell = sheet.cells.get((r, c2))
                if lbl_cell and val_cell:
                    labels.append(str(lbl_cell.value))
                    v = val_cell.cached if val_cell.formula else val_cell.value
                    if isinstance(v, (int, float)):
                        values.append(float(v))
        elif r2 - r1 == 1:
            for c in range(c1, c2 + 1):
                lbl_cell = sheet.cells.get((r1, c))
                val_cell = sheet.cells.get((r2, c))
                if lbl_cell and val_cell:
                    labels.append(str(lbl_cell.value))
                    v = val_cell.cached if val_cell.formula else val_cell.value
                    if isinstance(v, (int, float)):
                        values.append(float(v))
        if not values:
            messagebox.showwarning("Предупреждение", "Недостаточно данных для диаграммы")
            return
        try:
            import matplotlib
            matplotlib.use('TkAgg')
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            chart_win = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
            chart_win.title("Диаграмма")
            chart_win.geometry("600x400")
            fig = Figure(figsize=(6, 4), dpi=100)
            ax = fig.add_subplot(111)
            ax.bar(range(len(values)), values, color='#3399ff')
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=45, ha='right')
            ax.set_ylabel("Значение")
            fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=chart_win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        except ImportError:
            chart_win = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
            chart_win.title("Диаграмма")
            chart_win.geometry("600x400")
            cv = tk.Canvas(chart_win, bg="white")
            cv.pack(fill=tk.BOTH, expand=True)
            chart_win.update()
            w = cv.winfo_width()
            h = cv.winfo_height()
            margin = 40
            max_val = max(values) if values else 1
            bar_w = (w - 2 * margin) / len(values)
            for i, v in enumerate(values):
                x1 = margin + i * bar_w + 2
                x2 = x1 + bar_w - 4
                bar_h = (v / max_val) * (h - 2 * margin)
                y1 = h - margin - bar_h
                y2 = h - margin
                cv.create_rectangle(x1, y1, x2, y2, fill='#3399ff', outline='black')
                cv.create_text((x1 + x2) / 2, h - margin + 15, text=labels[i] if i < len(labels) else "",
                               font=("Calibri", 8), angle=45)
                cv.create_text((x1 + x2) / 2, y1 - 10, text=f"{v:g}", font=("Calibri", 8))

    def _auto_sum(self):
        row, col = self.active_cell
        sheet = self.workbook.active_sheet
        start_row = row - 1
        while start_row >= 0:
            cell = sheet.cells.get((start_row, col))
            if not cell or cell.value == "":
                break
            start_row -= 1
        start_row += 1
        if start_row < row:
            formula = f"=SUM({cell_ref_to_str(start_row, col)}:{cell_ref_to_str(row - 1, col)})"
            self._start_edit(formula)
            self._commit_edit()
        else:
            start_col = col - 1
            while start_col >= 0:
                cell = sheet.cells.get((row, start_col))
                if not cell or cell.value == "":
                    break
                start_col -= 1
            start_col += 1
            if start_col < col:
                formula = f"=SUM({cell_ref_to_str(row, start_col)}:{cell_ref_to_str(row, col - 1)})"
                self._start_edit(formula)
                self._commit_edit()

    def _insert_row(self, at_row: int):
        sheet = self.workbook.active_sheet
        new_cells = {}
        for (r, c), cell in sheet.cells.items():
            if r >= at_row:
                new_cells[(r + 1, c)] = cell
            else:
                new_cells[(r, c)] = cell
        sheet.cells = new_cells
        self._recalculate_all()
        self._redraw()

    def _insert_col(self, at_col: int):
        sheet = self.workbook.active_sheet
        new_cells = {}
        for (r, c), cell in sheet.cells.items():
            if c >= at_col:
                new_cells[(r, c + 1)] = cell
            else:
                new_cells[(r, c)] = cell
        sheet.cells = new_cells
        self._recalculate_all()
        self._redraw()

    def _delete_row(self, row: int):
        sheet = self.workbook.active_sheet
        new_cells = {}
        for (r, c), cell in sheet.cells.items():
            if r == row:
                continue
            elif r > row:
                new_cells[(r - 1, c)] = cell
            else:
                new_cells[(r, c)] = cell
        sheet.cells = new_cells
        self._recalculate_all()
        self._redraw()

    def _delete_col(self, col: int):
        sheet = self.workbook.active_sheet
        new_cells = {}
        for (r, c), cell in sheet.cells.items():
            if c == col:
                continue
            elif c > col:
                new_cells[(r, c - 1)] = cell
            else:
                new_cells[(r, c)] = cell
        sheet.cells = new_cells
        self._recalculate_all()
        self._redraw()

    def _auto_fit_column(self, col: int):
        sheet = self.workbook.active_sheet
        max_width = 50
        for r in range(Sheet.MAX_ROWS):
            cell = sheet.cells.get((r, col))
            if cell and cell.value != "":
                display = self._get_display_value(cell)
                w = len(display) * 8 + 16
                max_width = max(max_width, w)
        sheet.col_widths[col] = min(max_width, 400)
        self._redraw()

    def _set_col_width(self, col: int):
        current = self.workbook.active_sheet.get_col_width(col)
        new = simpledialog.askinteger("Ширина столбца", "Ширина (пиксели):", initialvalue=current)
        if new and new > 0:
            self.workbook.active_sheet.col_widths[col] = new
            self._redraw()

    def _set_row_height(self, row: int):
        current = self.workbook.active_sheet.get_row_height(row)
        new = simpledialog.askinteger("Высота строки", "Высота (пиксели):", initialvalue=current)
        if new and new > 0:
            self.workbook.active_sheet.row_heights[row] = new
            self._redraw()

    def _perform_autofill(self):
        if not self.autofill_start or not self.autofill_end:
            return
        r1, c1 = self.autofill_start
        r2, c2 = self.autofill_end
        if r2 > r1:
            for r in range(r1 + 1, r2 + 1):
                self._fill_cell_from_pattern(r, c1, r1, c1, r - r1)
        elif r2 < r1:
            for r in range(r1 - 1, r2 - 1, -1):
                self._fill_cell_from_pattern(r, c1, r1, c1, r1 - r)
        elif c2 > c1:
            for c in range(c1 + 1, c2 + 1):
                self._fill_cell_from_pattern(r1, c, r1, c1, c - c1)
        elif c2 < c1:
            for c in range(c1 - 1, c2 - 1, -1):
                self._fill_cell_from_pattern(r1, c, r1, c1, c1 - c)
        self.selection_start = self.autofill_start
        self.selection_end = self.autofill_end
        self._recalculate_all()
        self._redraw()

    def _fill_cell_from_pattern(self, tr: int, tc: int, sr: int, sc: int, offset: int):
        sheet = self.workbook.active_sheet
        source = sheet.cells.get((sr, sc))
        if not source:
            return
        target = sheet.get_cell(tr, tc)
        target.font_family = source.font_family
        target.font_size = source.font_size
        target.bold = source.bold
        target.italic = source.italic
        target.fg_color = source.fg_color
        target.bg_color = source.bg_color
        target.align = source.align
        target.border = source.border
        val = source.value
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            target.value = val + offset
        elif isinstance(val, str):
            m = re.match(r'^(.*?)(\d+)$', val)
            if m:
                prefix, num = m.groups()
                target.value = f"{prefix}{int(num) + offset}"
            else:
                target.value = val
        else:
            target.value = val
        target.formula = None
        target.cached = None

    def _find(self):
        dialog = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
        dialog.title("Найти")
        dialog.geometry("350x120")
        dialog.transient(self.root)
        tk.Label(dialog, text="Найти:").grid(row=0, column=0, padx=5, pady=5)
        entry = tk.Entry(dialog, width=30)
        entry.grid(row=0, column=1, padx=5, pady=5)
        entry.focus_set()
        result_var = tk.StringVar()
        tk.Label(dialog, textvariable=result_var).grid(row=2, column=0, columnspan=2, pady=5)

        def find_next():
            text = entry.get()
            if not text:
                return
            sheet = self.workbook.active_sheet
            start_r, start_c = self.active_cell
            for r in range(start_r, Sheet.MAX_ROWS):
                for c in range(start_c if r == start_r else 0, Sheet.MAX_COLS):
                    cell = sheet.cells.get((r, c))
                    if cell and text.lower() in str(cell.value).lower():
                        self.active_cell = (r, c)
                        self.selection_start = (r, c)
                        self.selection_end = (r, c)
                        self._scroll_to_cell(r, c)
                        self._update_name_box()
                        self._update_formula_bar()
                        self._redraw()
                        result_var.set(f"Найдено: {cell_ref_to_str(r, c)}")
                        return
            result_var.set("Не найдено")

        tk.Button(dialog, text="Найти далее", command=find_next).grid(row=1, column=1, padx=5, pady=5)

    def _replace(self):
        dialog = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
        dialog.title("Заменить")
        dialog.geometry("400x180")
        dialog.transient(self.root)
        tk.Label(dialog, text="Найти:").grid(row=0, column=0, padx=5, pady=5)
        find_entry = tk.Entry(dialog, width=30)
        find_entry.grid(row=0, column=1, padx=5, pady=5)
        find_entry.focus_set()
        tk.Label(dialog, text="Заменить на:").grid(row=1, column=0, padx=5, pady=5)
        replace_entry = tk.Entry(dialog, width=30)
        replace_entry.grid(row=1, column=1, padx=5, pady=5)
        count_var = tk.StringVar(value="")
        tk.Label(dialog, textvariable=count_var).grid(row=3, column=0, columnspan=2, pady=5)

        def replace_all():
            find_text = find_entry.get()
            replace_text = replace_entry.get()
            if not find_text:
                return
            sheet = self.workbook.active_sheet
            count = 0
            for (r, c), cell in list(sheet.cells.items()):
                if isinstance(cell.value, str) and find_text in cell.value:
                    cell.value = cell.value.replace(find_text, replace_text)
                    cell.cached = None
                    count += 1
            self._recalculate_all()
            self._redraw()
            count_var.set(f"Заменено: {count}")

        tk.Button(dialog, text="Заменить всё", command=replace_all).grid(row=2, column=1, padx=5, pady=5)

    def _set_zoom(self, zoom: float):
        self.zoom = max(0.5, min(2.0, zoom))
        self._redraw()
        self._update_status(f"Масштаб: {int(self.zoom * 100)}%")

    def _toggle_grid(self):
        self.show_grid = not self.show_grid
        self._redraw()

    def _show_about(self):
        messagebox.showinfo("О программе",
                            "PyExcel v1.0\n"
                            "Табличный процессор на Python\n"
                            "Поддерживает формулы, форматирование,\n"
                            "графики, множественные листы.\n"
                            "© 2026")

    def _show_functions_help(self):
        dialog = tk.Toplevel(getattr(self.root, "_afold_host", self.root))
        dialog.title("Поддерживаемые функции")
        dialog.geometry("600x500")
        text = tk.Text(dialog, wrap=tk.WORD, font=("Calibri", 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        help_text = """ПОДДЕРЖИВАЕМЫЕ ФУНКЦИИ
========================
МАТЕМАТИЧЕСКИЕ:
SUM(диапазон) — сумма
AVERAGE(диапазон) — среднее
MIN(диапазон) — минимум
MAX(диапазон) — максимум
COUNT(диапазон) — количество чисел
COUNTA(диапазон) — количество непустых
PRODUCT(диапазон) — произведение
MEDIAN(числа) — медиана
STDEV(числа) — стандартное отклонение
ABS(число) — модуль
ROUND(число; разряды) — округление
SQRT(число) — квадратный корень
POWER(основание; показатель) — степень
MOD(число; делитель) — остаток
PI() — число π
SIN/COS/TAN(угол) — тригонометрия
LOG(число; основание) — логарифм
LN(число) — натуральный логарифм
EXP(число) — экспонента
RAND() — случайное [0;1)
RANDBETWEEN(мин; макс) — случайное целое
ЛОГИЧЕСКИЕ:
IF(условие; да; нет) — условие
AND(условия) — логическое И
OR(условия) — логическое ИЛИ
NOT(условие) — отрицание
IFERROR(значение; ошибка) — обработка ошибок
ISBLANK/ISNUMBER/ISTEXT/ISERROR — проверки
ТЕКСТОВЫЕ:
CONCAT/CONCATENATE(тексты) — сцепка
LEN(текст) — длина
UPPER/LOWER/PROPER(текст) — регистр
TRIM(текст) — убрать пробелы
LEFT/RIGHT/MID(текст; кол-во) — подстроки
FIND(что; где; [начало]) — поиск
SUBSTITUTE(текст; старое; новое) — замена
REPT(текст; число) — повтор
EXACT(текст1; текст2) — точное сравнение
ДАТА И ВРЕМЯ:
NOW() — текущие дата и время
TODAY() — текущая дата
DATE(год; месяц; день) — дата
YEAR/MONTH/DAY(дата) — компоненты
HOUR/MINUTE/SECOND(дата) — время
ПОИСК И ССЫЛКИ:
VLOOKUP(ключ; таблица; столбец) — вертикальный поиск
HLOOKUP(ключ; таблица; строка) — горизонтальный поиск
INDEX(массив; номер) — элемент по индексу
MATCH(значение; массив) — позиция значения
CHOOSE(индекс; значения) — выбор по индексу
УСЛОВНЫЕ:
COUNTIF(диапазон; критерий) — подсчёт по условию
SUMIF(диапазон; критерий; [сумм_диапазон]) — сумма по условию
AVERAGEIF(диапазон; критерий; [диапазон]) — среднее по условию
ПРИМЕРЫ:
=SUM(A1:A10)
=AVERAGE(B1:B5)
=IF(A1>10; "Много"; "Мало")
=VLOOKUP(D1; A1:B10; 2)
=CONCAT(A1; " "; B1)
=ROUND(3.14159; 2)
=TODAY()
"""
        text.insert(1.0, help_text)
        text.config(state=tk.DISABLED)

# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================
def main():
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except Exception:
        pass
    app = SpreadsheetApp(root)

    def on_closing():
        if messagebox.askokcancel("Выход", "Выйти из программы?"):
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()