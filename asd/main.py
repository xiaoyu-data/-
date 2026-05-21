from agent import OfficeAgent


def main():
    print(r"""
╔══════════════════════════════════════════════════════╗
║           OfficeAgent - 办公自动化 Agent             ║
║                                                      ║
║  支持操作:                                           ║
║    Word  - 创建、读取、替换文字、提取表格             ║
║    Excel - 创建、读取、写入、数据分析                 ║
║    PDF   - 提取文本、提取表格、获取信息               ║
║    文件  - 列表、读写、创建目录、复制删除             ║
║                                                      ║
║  示例命令:                                           ║
║    创建一个 Word 文档                                 ║
║    读取 Excel 文件 report.xlsx                        ║
║    提取 PDF 中的文本 document.pdf                     ║
║    列出当前目录的文件                                 ║
║    帮我把 data.docx 中的 A 替换成 B                   ║
║    分析 data.xlsx                                     ║
║    退出                                               ║
╚══════════════════════════════════════════════════════╝
""")

    agent = OfficeAgent(workspace=r"c:\Users\熊凯鹏\Desktop\asd")

    while True:
        try:
            task = input("\n[Agent] 请输入任务 > ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n再见！")
            break

        if not task:
            continue

        if task in ("退出", "exit", "quit", "q"):
            print("再见！")
            break

        result = agent.run(task)
        print(f"\n{'─' * 60}")
        print(f"[结果]\n{result}")
        print(f"{'─' * 60}")


if __name__ == "__main__":
    main()
