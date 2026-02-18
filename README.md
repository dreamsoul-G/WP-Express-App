# WP Express (Document Processing) Software WP快通（文档处理）软件
A lightweight Python desktop tool for Word (.docx) and PDF documents processing (Windows only),referred to as WP Express for short.
> 轻量级 Python 桌面工具，专注于 Word (.docx) 和 PDF 文档处理（仅支持 Windows），简称WP快通。

# WP Express (WP快通)
A lightweight Python desktop tool for Word (.docx) and PDF documents processing (Windows only).
> 轻量级 Python 桌面工具，专注于 Word (.docx) 和 PDF 文档处理（仅支持 Windows）。


## 📖 About (关于项目)
### Motivation (开发初衷)
As a college student from China, this tool was born out of a real-life hassle:  
When completing a school paper assignment, I needed to submit the final version in PDF format, and the cover page had to be signed and then photographed for upload. The main body of the paper was written in Word, but I encountered two difficulties during the process:
1. Merging the cover page (converted from photos) with the main thesis document required VIP access on mainstream tools like WPS;  
2. Two-way conversion between Word and PDF; most online tools/software also have a paywall.

Frustrated by these unnecessary costs for basic document tasks, I decided to build a simple, free tool to solve these pain points—WP Express was born. Over time, it has evolved to cover more basic document needs for myself and my peers.


> 作为一名中国大学生，这个工具的诞生源于一次真实地困扰：
> 某次完成学校论文作业时，我需要提交 PDF 格式的终稿，而封面需签名后再拍照上传。论文主体用 Word 编写，但过程中遇到了两个难题：
> 1. 将拍照后的封面（转为 Word 格式）与论文正文合并，主流工具（如 WPS）需要开通 VIP 才能实现；  
> 2. Word 与 PDF 双向转换，多数线上工具/软件也设有付费门槛。 

> 为基础的文档处理需求支付不必要的费用让我很困扰，于是我决定做一个简单、免费的工具来解决这些问题——WP Express 就此诞生。如今它已逐步完善，能满足我和身边同学更多的基础文档处理需求。

### Core Features (核心功能)
- **Merge & Split**: Merge multiple Word/PDF files into one, or split a single document into two files;
- **Bidirectional Conversion**: Seamlessly convert between Word (.docx) and PDF formats.

> - **拆分与合并**：将多个 Word/PDF 文件合并为一个，或把单个文档拆分为两个文件；
> - **双向转换**：在 Word (.docx) 和 PDF 格式之间无缝转换。

## 🌐 Language (语言)  
- Current version: Simplified Chinese.
> - 当前版本：简体中文。

## 🚀 Development Plan (开发规划)

1. Refactor GUI with PySide6, build local lightweight SQLite database, and implement multithreaded development;
2. Add more free document processing features to cover daily academic needs;
3. Develop paid closed-source features (e.g., batch multi-file conversion) for advanced use cases;
4. Optimize page interaction and thread processing logic (improve stability and user experience).


> 1. 基于 PySide6 重构 GUI，搭建本地轻量级 SQLite 数据库，实现多线程开发；
> 2. 新增更多免费文档处理功能，覆盖日常学业需求；
> 3. 开发付费闭源功能（如批量多文件转换），满足进阶使用场景；
> 4. 优化页面交互和线程处理逻辑（提升稳定性与用户体验）。

## 🧑‍💻 Developer (开发者)
- GitHub ID: dreamsoul-G;
- Identity: College Student from China;
- Tech Stack: Python, Tkinter/PySide6, PyInstaller, SQLite, python-docx, PyPDF2.

> - GitHub 账户名：dreamsoul-G；
> - 身份：中国大学生；
> - 技术栈：Python、tkinter、PyInstaller、json、python-docx、pypdf。

## 📦 Dependencies (依赖库)
All used third-party libraries are open-source and commercially usable (with specified version requirements for stability):
- python-docx>=1.2.0: For Word (.docx) file processing;
- pypdf>=6.5.0: For PDF file parsing and manipulation;
- reportlab>=4.4.9: For PDF file generation;
- Pillow>=10.4.0: For image processing;
- pywin32>=311 (Windows only): For Windows-specific conversion features.

> 所有第三方依赖库均为开源且可商用（标注版本号保证稳定性）：
> - python-docx>=1.2.0：处理 Word (.docx) 文件；
> - pypdf>=6.5.0：解析和操作 PDF 文件；
> - reportlab>=4.4.9：生成 PDF 文件；
> - Pillow>=10.4.0：处理图片；
> - pywin32>=311（仅Windows）：实现Windows专属的转换功能。

## 📄 License (许可证)
This project is licensed under the Apache License 2.0 - see the LICENSE file for details.
> 本项目基于 Apache License 2.0 协议开源（核心免费功能永久开源，付费功能闭源）。

## 📥 Download (下载)
Get the latest Windows executable (.exe) from the [Releases](https://github.com/dreamsoul-G/WP-Express-App/releases) page.
> 从 [Releases](https://github.com/dreamsoul-G/WP-Express-App/releases) 页面下载最新 Windows 可执行文件（.exe）。

## 💖 A Note to Users (致用户)
I hope WP Express can solve your daily document processing troubles just like it did for me. If this tool helps you, your support (star/fork/share) would mean a lot to me as a student developer!
> 希望 WP 快通能像帮助我一样，解决你日常文档处理的烦恼。如果这个工具对你有用，作为一名学生开发者，你的支持（点赞/复刻/分享）对我来说意义非凡！
