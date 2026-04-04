## 英文单词处理工具 v1.5
=====================
### 功能
从txt、markdown、PDF或EPUB文档中提取去重后的英文单词排序，并统计总词数和某单词出现的次数，可选输出为txt或excel（包括CSV）。
---
使用方法：
1. 双击运行"word_count_gui.exe"或者python环境下双击“word_count_gui.py”
2. 选择输入文件（TXT、Markdown、PDF 或 EPUB）
3. 选择输出目录
4. 右边窗口包含可自定义添加的排除词
5. 点击"开始处理"
<img width="1920" height="1020" alt="v1 5" src="https://github.com/user-attachments/assets/bd563554-1ec5-490b-8227-5019e574b161" />

依赖说明：
- 如果只在 Python 环境运行.py 文件，想要针对 PDF 操作，需要安装 pip install pdfplumber
- 如果只在 Python 环境运行.py 文件，想要针对 EPUB 操作，需要安装 pip install ebooklib
- 首次运行可能需要安装 Visual C++ 运行库
- 如被杀毒软件拦截，请添加信任
