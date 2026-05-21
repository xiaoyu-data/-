import re
import os
from pathlib import Path
from typing import Any, Callable
from tools import FileTools, WordTools, ExcelTools, PDFTools


class OfficeAgent:
    def __init__(self, workspace: str = "."):
        self.workspace = Path(workspace).resolve()
        self.tools: dict[str, Any] = {}
        self.steps: list[dict] = []
        self.results: list[str] = []

        self._register_default_tools()

    def _register_default_tools(self):
        self.register_tool(FileTools())
        self.register_tool(WordTools())
        self.register_tool(ExcelTools())
        self.register_tool(PDFTools())

    def register_tool(self, tool):
        self.tools[tool.name] = tool
        print(f"[Agent] 已注册工具: {tool.name} - {tool.description}")

    def run(self, task: str) -> str:
        self.steps = []
        self.results = []
        print(f"\n{'='*60}")
        print(f"[Agent] 收到任务: {task}")
        print(f"{'='*60}")

        plan = self._plan(task)

        if not plan:
            return self._fallback(task)

        for i, step in enumerate(plan, 1):
            print(f"\n[Step {i}/{len(plan)}] {step['description']}")
            result = self._execute_step(step)
            self.results.append(result)
            print(f"  结果: {result}")

        return self._summarize()

    @staticmethod
    def _has_any_word(text: str, words: list[str]) -> bool:
        return any(w in text for w in words)

    @staticmethod
    def _has_all_words(text: str, word_groups: list[list[str]]) -> bool:
        return all(any(w in text for w in group) for group in word_groups)

    def _plan(self, task: str) -> list[dict]:
        task_lower = task.lower()
        filepath = self._extract_filepath(task)

        if self._has_all_words(task_lower, [["创建", "生成", "新建", "写"], ["word", "docx", "文档"]]) and \
           not self._has_any_word(task_lower, ["excel", "xlsx", "表格", "pdf"]):
            title = self._extract_title(task)
            content = self._extract_content(task)
            path = filepath or str(self.workspace / "output.docx")
            return [{
                "tool": "WordTools",
                "action": "create",
                "description": f"创建 Word 文档: {path}",
                "kwargs": {"filepath": path, "content": content, "title": title or "自动生成文档"},
            }]

        if self._has_all_words(task_lower, [["读取", "查看", "打开", "阅读", "读"], ["word", "docx", "文档"]]) and filepath:
            return [{
                "tool": "WordTools",
                "action": "read",
                "description": f"读取 Word 文档: {filepath}",
                "kwargs": {"filepath": filepath},
            }]

        if self._has_any_word(task_lower, ["替换", "取代", "换成"]) and \
           self._has_any_word(task_lower, ["word", "docx", "文档"]):
            old_text, new_text = self._extract_replace_pair(task)
            output = self._extract_output(task)
            return [{
                "tool": "WordTools",
                "action": "replace",
                "description": "替换 Word 中的文字",
                "kwargs": {"filepath": filepath, "old_text": old_text, "new_text": new_text, "output": output or ""},
            }]

        if self._has_all_words(task_lower, [["提取", "导出", "获取"], ["表格"]]) and \
           self._has_any_word(task_lower, ["word", "docx", "文档"]) and \
           not self._has_any_word(task_lower, ["pdf"]):
            return [{
                "tool": "WordTools",
                "action": "tables",
                "description": "提取 Word 文档中的表格",
                "kwargs": {"filepath": filepath},
            }]

        if self._has_all_words(task_lower, [["创建", "生成", "新建", "写"], ["excel", "xlsx", "表格", "工作簿"]]) and \
           not self._has_any_word(task_lower, ["word", "docx", "文档", "pdf"]):
            path = filepath or str(self.workspace / "output.xlsx")
            data = self._extract_table_data(task)
            return [{
                "tool": "ExcelTools",
                "action": "create",
                "description": f"创建 Excel 工作簿: {path}",
                "kwargs": {"filepath": path, "sheet_data": {"data": data}},
            }]

        if self._has_all_words(task_lower, [["读取", "查看", "打开", "阅读", "读"], ["excel", "xlsx", "表格"]]) and filepath:
            return [{
                "tool": "ExcelTools",
                "action": "read",
                "description": f"读取 Excel 文件: {filepath}",
                "kwargs": {"filepath": filepath},
            }]

        if self._has_all_words(task_lower, [["分析", "统计", "汇总"], ["excel", "xlsx", "表格", "数据"]]) and filepath:
            return [{
                "tool": "ExcelTools",
                "action": "analyze",
                "description": f"分析 Excel 数据: {filepath}",
                "kwargs": {"filepath": filepath},
            }]

        if self._has_all_words(task_lower, [["写入", "追加", "添加"], ["excel", "xlsx", "表格"]]) and filepath:
            data = self._extract_table_data(task)
            sheet = self._extract_sheet_name(task)
            return [{
                "tool": "ExcelTools",
                "action": "write",
                "description": "向 Excel 写入数据",
                "kwargs": {"filepath": filepath, "data": data, "sheet_name": sheet or "Sheet1"},
            }]

        if self._has_any_word(task_lower, ["pdf"]) and \
           self._has_any_word(task_lower, ["提取", "读取", "查看", "打开", "阅读", "读"]) and \
           not self._has_any_word(task_lower, ["表格"]):
            return [{
                "tool": "PDFTools",
                "action": "extract_text",
                "description": f"提取 PDF 文本: {filepath}",
                "kwargs": {"filepath": filepath},
            }]

        if self._has_any_word(task_lower, ["pdf"]) and \
           self._has_any_word(task_lower, ["表格"]):
            return [{
                "tool": "PDFTools",
                "action": "extract_tables",
                "description": f"提取 PDF 表格: {filepath}",
                "kwargs": {"filepath": filepath},
            }]

        if self._has_any_word(task_lower, ["pdf"]) and \
           self._has_any_word(task_lower, ["信息", "属性", "详情", "元数据"]):
            return [{
                "tool": "PDFTools",
                "action": "get_info",
                "description": f"获取 PDF 信息: {filepath}",
                "kwargs": {"filepath": filepath},
            }]

        if self._has_any_word(task_lower, ["列出", "查看", "显示", "ls"]) and \
           self._has_any_word(task_lower, ["文件", "目录", "文件夹"]):
            path = filepath or str(self.workspace)
            return [{
                "tool": "FileTools",
                "action": "list",
                "description": f"列出文件: {path}",
                "kwargs": {"path": path},
            }]

        if self._has_all_words(task_lower, [["读取", "查看", "打开", "读"], ["文件"]]) and filepath:
            return [{
                "tool": "FileTools",
                "action": "read",
                "description": f"读取文件: {filepath}",
                "kwargs": {"filepath": filepath},
            }]

        if self._has_all_words(task_lower, [["创建", "新建", "mkdir"], ["目录", "文件夹"]]):
            path = filepath or str(self.workspace / "新文件夹")
            return [{
                "tool": "FileTools",
                "action": "mkdir",
                "description": f"创建目录: {path}",
                "kwargs": {"path": path},
            }]

        return []

    def _execute_step(self, step: dict) -> str:
        tool_name = step["tool"]
        action = step["action"]
        kwargs = step.get("kwargs", {})

        tool = self.tools.get(tool_name)
        if not tool:
            return f"未找到工具: {tool_name}"

        try:
            return tool.run(action=action, **kwargs)
        except Exception as e:
            return f"执行失败 ({tool_name}.{action}): {e}"

    def _summarize(self) -> str:
        if not self.results:
            return "未能识别任务。请尝试更明确地描述您的需求，例如：\n" \
                   '  - "创建一个 Word 文档"\n' \
                   '  - "读取 Excel 文件 data.xlsx"\n' \
                   '  - "提取 PDF 中的文本 report.pdf"\n' \
                   '  - "列出当前目录的文件"'
        return "\n".join(f"[结果 {i+1}] {r}" for i, r in enumerate(self.results))

    def _fallback(self, task: str) -> str:
        print("[Agent] 无法自动识别任务类型，尝试智能推断...")
        filepath = self._extract_filepath(task)

        if not filepath:
            return self._summarize()

        ext = Path(filepath).suffix.lower()

        if ext in (".docx", ".doc"):
            step = {"tool": "WordTools", "action": "read", "description": f"读取: {filepath}", "kwargs": {"filepath": filepath}}
        elif ext in (".xlsx", ".xls"):
            step = {"tool": "ExcelTools", "action": "read", "description": f"读取: {filepath}", "kwargs": {"filepath": filepath}}
        elif ext == ".pdf":
            step = {"tool": "PDFTools", "action": "extract_text", "description": f"提取: {filepath}", "kwargs": {"filepath": filepath}}
        else:
            step = {"tool": "FileTools", "action": "read", "description": f"读取: {filepath}", "kwargs": {"filepath": filepath}}

        result = self._execute_step(step)
        self.results.append(result)
        print(f"  结果: {result}")
        return self._summarize()

    def _extract_filepath(self, text: str) -> str:
        patterns = [
            r'(?:文件|路径|文档|表格)[：:]\s*(\S+)',
            r'[\'\"](\S+\.(?:docx?|xlsx?|pdf|txt|csv|json))[\'\"]',
            r'(\S+\.(?:docx?|xlsx?|pdf|txt|csv|json))',
        ]
        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                path = m.group(1)
                full = self.workspace / path
                if full.exists():
                    return str(full)
                return path
        return ""

    def _extract_title(self, text: str) -> str:
        m = re.search(r'标题[：:]\s*(.+?)(?:[，,。\n]|$)', text)
        if m:
            return m.group(1).strip()
        return ""

    def _extract_content(self, text: str) -> str:
        m = re.search(r'内容[：:]\s*(.+?)(?:[。\n]|$)', text)
        if m:
            return m.group(1).strip()
        return "这是一个由 OfficeAgent 自动生成的文档。\n文档内容可以根据您的需求进行自定义。"

    def _extract_replace_pair(self, text: str) -> tuple:
        old_text = ""
        new_text = ""
        m = re.search(r'把[「「"\']?(.+?)[」」"\']?\s*(?:替换|改|换成)[为：:]?\s*[「「"\']?(.+?)[」」"\']?', text)
        if m:
            old_text = m.group(1).strip()
            new_text = m.group(2).strip()
        return old_text, new_text

    def _extract_output(self, text: str) -> str:
        m = re.search(r'保存[到至]?[：:]?\s*(\S+)', text)
        if m:
            return m.group(1)
        return ""

    def _extract_table_data(self, text: str) -> list:
        data = []
        json_match = re.search(r'\[\[.+?\]\]', text)
        if json_match:
            try:
                import json
                data = json.loads(json_match.group())
                return data
            except json.JSONDecodeError:
                pass
        return data or [
            ["姓名", "部门", "分数"],
            ["张三", "技术部", "92"],
            ["李四", "市场部", "85"],
        ]

    def _extract_sheet_name(self, text: str) -> str:
        m = re.search(r'工作表[：:]?\s*(\S+)', text)
        if m:
            return m.group(1)
        m = re.search(r'[Ss]heet[：:]?\s*(\S+)', text)
        if m:
            return m.group(1)
        return ""


def demo():
    agent = OfficeAgent(workspace=r"c:\Users\熊凯鹏\Desktop\asd")

    print("\n" + "=" * 60)
    print("  OfficeAgent - 办公自动化 Agent 演示")
    print("=" * 60)

    tasks = [
        "创建一个 Word 文档，标题：会议纪要，内容：今天讨论了Q2的项目计划",
        "列出当前目录的所有文件",
    ]

    for task in tasks:
        result = agent.run(task)
        print(f"\n{'='*60}")
        print(f"最终结果:\n{result}")
        print(f"{'='*60}")

    print("\n[Agent] 你可以通过以下方式与我交互：")
    print('  agent.run("创建一个 Word 文档")')
    print('  agent.run("读取 Excel 文件 data.xlsx")')
    print('  agent.run("提取 PDF 中的文本 report.pdf")')
    print('  agent.run("列出当前目录的文件")')
    print('  agent.run("帮我把 data.docx 中的 旧词 替换成 新词")')


if __name__ == "__main__":
    demo()
