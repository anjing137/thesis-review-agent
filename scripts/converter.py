#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档转换模块 - 使用Pandoc将Word文档转换为Markdown:

支持Windows和Mac平台，自动检测Pandoc路径
支持 .docx 和 .doc 格式（.doc 会先用 Word 转换为 .docx）

[DEPRECATED] 此模块已废弃，仅用于 .md 旧路径的向后兼容。
新流程（.docx 直接分析）请使用 xml_analyzer 模块，无需此转换。
"""

import os
import subprocess
import sys
import tempfile
from pathlib import Path


class Converter:
    """文档转换器"""

    # Pandoc路径配置
    PANDOC_PATHS = {
        'win32': [
            r"D:\download\Programs\pandoc\pandoc.exe",
            r"C:\Program Files\Pandoc\pandoc.exe",
            r"C:\Program Files (x86)\Pandoc\pandoc.exe",
        ],
        'darwin': [
            '/usr/local/bin/pandoc',
            '/opt/homebrew/bin/pandoc',
            '/Applications/Pandoc.app/Contents/MacOS/pandoc',
        ],
        'linux': [
            '/usr/bin/pandoc',
            '/usr/local/bin/pandoc',
        ]
    }

    def __init__(self, paper_path: str = None):
        """
        初始化转换器

        Args:
            paper_path: 论文文件路径（可选，后续再设置）
        """
        self.paper_path = paper_path
        self.pandoc_path = self._find_pandoc()

    def _find_pandoc(self) -> str:
        """查找Pandoc可执行文件路径"""
        platform = sys.platform

        # 首先检查环境变量
        env_path = os.environ.get('PANDOC_PATH')
        if env_path and os.path.exists(env_path):
            return env_path

        # 根据平台搜索预设路径
        paths = self.PANDOC_PATHS.get(platform, [])
        for path in paths:
            if os.path.exists(path):
                return path

        # 尝试在PATH中查找
        try:
            result = subprocess.run(
                ['pandoc', '--version'],
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                return 'pandoc'  # 在PATH中
        except FileNotFoundError:
            pass

        # 未找到
        return None

    def is_pandoc_available(self) -> bool:
        """检查Pandoc是否可用"""
        return self.pandoc_path is not None

    def _convert_doc_to_docx(self, doc_path: str) -> str:
        """
        将 .doc 转换为 .docx（使用 Word COM 接口）

        Args:
            doc_path: .doc 文件路径

        Returns:
            .docx 文件路径，失败返回 None
        """
        if sys.platform != 'win32':
            return None

        try:
            import win32com.client as win32

            # 创建临时目录
            temp_dir = tempfile.gettempdir()
            base_name = os.path.splitext(os.path.basename(doc_path))[0]
            docx_path = os.path.join(temp_dir, f"{base_name}.docx")

            # 使用 Word 转换
            word = win32.Dispatch("Word.Application")
            word.Visible = False

            try:
                doc = word.Documents.Open(os.path.abspath(doc_path))
                doc.SaveAs2(os.path.abspath(docx_path), FileFormat=16)  # 16 = wdFormatXMLDocument
                doc.Close()
            finally:
                word.Quit()

            return docx_path

        except ImportError:
            return None
        except Exception as e:
            print(f"DOC转DOCX失败: {e}")
            return None

    def convert(self, paper_path: str = None, output_dir: str = None) -> dict:
        """
        转换Word文档为Markdown

        Args:
            paper_path: 论文文件路径（如果初始化时未设置）
            output_dir: 输出目录（可选，默认与源文件同目录）

        Returns:
            转换结果字典：
            {
                'success': bool,
                'md_path': str,  # Markdown文件路径
                'error': str     # 错误信息（如果失败）
            }
        """
        if paper_path:
            self.paper_path = paper_path

        if not self.paper_path:
            return {
                'success': False,
                'md_path': None,
                'error': '未指定论文文件路径'
            }

        if not os.path.exists(self.paper_path):
            return {
                'success': False,
                'md_path': None,
                'error': f'文件不存在：{self.paper_path}'
            }

        if not self.pandoc_path:
            return {
                'success': False,
                'md_path': None,
                'error': '未找到Pandoc，请安装Pandoc或设置PANDOC_PATH环境变量'
            }

        # 检查是否是 .doc 格式，需要转换为 .docx
        original_path = self.paper_path
        is_doc = self.paper_path.lower().endswith('.doc')
        temp_docx_path = None

        if is_doc:
            temp_docx_path = self._convert_doc_to_docx(self.paper_path)
            if temp_docx_path:
                self.paper_path = temp_docx_path
            else:
                return {
                    'success': False,
                    'md_path': None,
                    'error': 'Pandoc不支持DOC格式，且无法调用Word进行转换。请手动将DOC文件另存为DOCX格式后再试。'
                }

        # 确定输出路径
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(original_path))[0]
            md_path = os.path.join(output_dir, f'{base_name}.md')
        else:
            md_path = original_path.replace('.docx', '.md').replace('.doc', '.md')

        # 执行转换
        try:
            cmd = [self.pandoc_path, self.paper_path, '-o', md_path]
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                errors='replace'
            )

            if result.returncode == 0:
                return {
                    'success': True,
                    'md_path': md_path,
                    'error': None
                }
            else:
                return {
                    'success': False,
                    'md_path': None,
                    'error': result.stderr or '转换失败'
                }

        except Exception as e:
            return {
                'success': False,
                'md_path': None,
                'error': str(e)
            }
        finally:
            # 清理临时文件
            if temp_docx_path and os.path.exists(temp_docx_path):
                try:
                    os.remove(temp_docx_path)
                except:
                    pass


def convert_paper(paper_path: str, output_dir: str = None) -> dict:
    """
    便捷函数：转换单篇论文

    Args:
        paper_path: 论文文件路径
        output_dir: 输出目录（可选）

    Returns:
        转换结果字典
    """
    converter = Converter(paper_path)
    return converter.convert(output_dir=output_dir)


if __name__ == '__main__':
    # 测试代码
    if len(sys.argv) > 1:
        paper_path = sys.argv[1]
        output_dir = sys.argv[2] if len(sys.argv) > 2 else None

        result = convert_paper(paper_path, output_dir)
        if result['success']:
            print(f"✅ 转换成功：{result['md_path']}")
        else:
            print(f"❌ 转换失败：{result['error']}")
            sys.exit(1)
    else:
        print("用法: python converter.py <论文路径> [输出目录]")
        sys.exit(1)