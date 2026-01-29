"""
Manual Dialog Module
Provides user manual and documentation
"""
import os
import sys
import re
import importlib
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                            QTextBrowser, QTabWidget, QWidget)
from PyQt5.QtCore import Qt
# Lazy import markdown to avoid py2app zipimporter issues
# markdown will be imported only when load_markdown_file is called


class ManualDialog(QDialog):
    """Manual Dialog showing user documentation"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("User Manual")
        self.setMinimumWidth(700)
        self.setMinimumHeight(600)
        
        # Get the base directory for markdown files
        if getattr(sys, 'frozen', False):
            # Running as compiled app (py2app)
            # sys.executable = Contents/MacOS/Slide Maker
            # Resources = Contents/Resources
            base_dir = os.path.dirname(os.path.dirname(sys.executable))  # Contents
            resources_dir = os.path.join(base_dir, "Resources")  # Contents/Resources
            self.docs_dir = os.path.join(resources_dir, "docs")  # Contents/Resources/docs
        else:
            # Running as script
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            self.docs_dir = os.path.join(base_dir, "docs")
        
        self.setup_ui()
    
    def _simple_markdown_to_html(self, text):
        """Simple markdown to HTML converter as fallback with improved support"""
        lines = text.split('\n')
        html_lines = []
        in_code_block = False
        code_block_content = []
        in_table = False
        in_list = False
        list_type = None  # 'ul' or 'ol'
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # Code blocks (```)
            if stripped.startswith('```'):
                if in_code_block:
                    # End code block
                    code_content = '\n'.join(code_block_content)
                    # Escape HTML in code
                    code_content = code_content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    html_lines.append(f'<pre style="background-color: #f5f5f5; padding: 10px; border-radius: 5px; overflow-x: auto; white-space: pre-wrap; font-family: monospace;"><code>{code_content}</code></pre>')
                    code_block_content = []
                    in_code_block = False
                else:
                    # Start code block
                    in_code_block = True
                i += 1
                continue
            
            if in_code_block:
                code_block_content.append(line)
                i += 1
                continue
            
            # End list if we were in one and this line doesn't continue it
            if in_list:
                is_list_item = bool(re.match(r'^[\s]*[-*+]\s+', line) or re.match(r'^[\s]*\d+\.\s+', line))
                if not is_list_item and stripped:
                    html_lines.append(f'</{list_type}>')
                    in_list = False
                    list_type = None
            
            # Tables (lines with |)
            if '|' in line and not stripped.startswith('#'):
                # Check if it's a table separator (contains dashes)
                if re.match(r'^\s*\|[\s\-:]+\|\s*$', line):
                    # Table separator, skip it
                    i += 1
                    continue
                
                # Parse table row
                cells = [cell.strip() for cell in line.split('|')]
                # Remove empty cells at start/end
                if cells and not cells[0]:
                    cells = cells[1:]
                if cells and not cells[-1]:
                    cells = cells[:-1]
                
                if cells:
                    if not in_table:
                        # Start table
                        html_lines.append('<table style="border-collapse: collapse; width: 100%; margin: 10px 0;">')
                        in_table = True
                        # First row is header
                        html_lines.append('<thead><tr>')
                        for cell in cells:
                            # Process inline formatting in cells
                            cell_html = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', cell)
                            cell_html = re.sub(r'\*(.*?)\*', r'<em>\1</em>', cell_html)
                            html_lines.append(f'<th style="border: 1px solid #ddd; padding: 8px; background-color: #f2f2f2; text-align: left;">{cell_html}</th>')
                        html_lines.append('</tr></thead><tbody>')
                    else:
                        # Data row
                        html_lines.append('<tr>')
                        for cell in cells:
                            # Process inline formatting in cells
                            cell_html = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', cell)
                            cell_html = re.sub(r'\*(.*?)\*', r'<em>\1</em>', cell_html)
                            html_lines.append(f'<td style="border: 1px solid #ddd; padding: 8px;">{cell_html}</td>')
                        html_lines.append('</tr>')
                i += 1
                continue
            else:
                # End table if we were in one
                if in_table:
                    html_lines.append('</tbody></table>')
                    in_table = False
            
            # Headers
            if stripped.startswith('### '):
                content = stripped[4:]
                # Process inline formatting
                content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', content)
                html_lines.append(f'<h3 style="margin-top: 20px; margin-bottom: 10px;">{content}</h3>')
            elif stripped.startswith('## '):
                content = stripped[3:]
                # Process inline formatting
                content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', content)
                html_lines.append(f'<h2 style="margin-top: 25px; margin-bottom: 15px;">{content}</h2>')
            elif stripped.startswith('# '):
                content = stripped[2:]
                # Process inline formatting
                content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', content)
                html_lines.append(f'<h1 style="margin-top: 30px; margin-bottom: 20px;">{content}</h1>')
            # Lists
            elif re.match(r'^[\s]*[-*+]\s+', line):
                # Unordered list
                if not in_list or list_type != 'ul':
                    if in_list:
                        html_lines.append(f'</{list_type}>')
                    html_lines.append('<ul style="margin: 10px 0; padding-left: 30px;">')
                    in_list = True
                    list_type = 'ul'
                
                indent = len(line) - len(line.lstrip())
                content = re.sub(r'^[\s]*[-*+]\s+', '', line)
                # Process inline formatting
                content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', content)
                content = re.sub(r'(?<!\*)\*([^*]+?)\*(?!\*)', r'<em>\1</em>', content)
                # Inline code
                content = re.sub(r'`([^`]+)`', r'<code style="background-color: #f5f5f5; padding: 2px 4px; border-radius: 3px; font-family: monospace;">\1</code>', content)
                html_lines.append(f'<li style="margin: 5px 0;">{content}</li>')
            elif re.match(r'^[\s]*\d+\.\s+', line):
                # Ordered list
                if not in_list or list_type != 'ol':
                    if in_list:
                        html_lines.append(f'</{list_type}>')
                    html_lines.append('<ol style="margin: 10px 0; padding-left: 30px;">')
                    in_list = True
                    list_type = 'ol'
                
                indent = len(line) - len(line.lstrip())
                content = re.sub(r'^[\s]*\d+\.\s+', '', line)
                # Process inline formatting
                content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', content)
                content = re.sub(r'(?<!\*)\*([^*]+?)\*(?!\*)', r'<em>\1</em>', content)
                # Inline code
                content = re.sub(r'`([^`]+)`', r'<code style="background-color: #f5f5f5; padding: 2px 4px; border-radius: 3px; font-family: monospace;">\1</code>', content)
                html_lines.append(f'<li style="margin: 5px 0;">{content}</li>')
            # Empty line
            elif not stripped:
                html_lines.append('<br>')
            # Regular paragraph
            else:
                # Process inline formatting
                content = line
                # Bold
                content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', content)
                # Italic (but not if it's part of bold)
                content = re.sub(r'(?<!\*)\*([^*]+?)\*(?!\*)', r'<em>\1</em>', content)
                # Inline code
                content = re.sub(r'`([^`]+)`', r'<code style="background-color: #f5f5f5; padding: 2px 4px; border-radius: 3px; font-family: monospace;">\1</code>', content)
                # Links
                content = re.sub(r'\[([^\]]+)\]\(([^\)]+)\)', r'<a href="\2">\1</a>', content)
                html_lines.append(f'<p style="margin: 10px 0; line-height: 1.6;">{content}</p>')
            
            i += 1
        
        # Close any open blocks
        if in_code_block:
            code_content = '\n'.join(code_block_content)
            # Escape HTML in code
            code_content = code_content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            html_lines.append(f'<pre style="background-color: #f5f5f5; padding: 10px; border-radius: 5px; overflow-x: auto; white-space: pre-wrap; font-family: monospace;"><code>{code_content}</code></pre>')
        if in_table:
            html_lines.append('</tbody></table>')
        if in_list:
            html_lines.append(f'</{list_type}>')
        
        return '\n'.join(html_lines)
    
    def load_markdown_file(self, filename):
        """Load and convert markdown file to HTML, removing all font-family declarations"""
        filepath = os.path.join(self.docs_dir, filename)
        
        if not os.path.exists(filepath):
            return f"<p>Error: Documentation file not found: {filepath}</p>"
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                markdown_content = f.read()
            
            # Try to use markdown library, but fallback to simple parser if it fails
            html = None
            try:
                # Try importing markdown module
                if 'markdown' not in sys.modules:
                    # Direct import without using importlib to avoid zipimporter issues
                    import markdown
                else:
                    markdown = sys.modules['markdown']
                
                # Try to use extensions for better markdown support
                try:
                    # Try with extensions for tables and code highlighting
                    extensions = ['extra', 'codehilite', 'tables']
                    html = markdown.markdown(markdown_content, extensions=extensions)
                except:
                    # If extensions fail, try with just tables
                    try:
                        extensions = ['extra', 'tables']
                        html = markdown.markdown(markdown_content, extensions=extensions)
                    except:
                        # If that fails, use basic markdown
                        html = markdown.markdown(markdown_content)
            except Exception as e:
                # If markdown library fails, use simple fallback parser
                html = self._simple_markdown_to_html(markdown_content)
            
            # Remove ALL font-family declarations from the generated HTML
            # This prevents Qt font warnings by removing any font references
            html = re.sub(r'font-family\s*:\s*[^;"]+;?', '', html, flags=re.IGNORECASE)
            html = re.sub(r'font-family\s*:\s*[^}]+', '', html, flags=re.IGNORECASE)
            
            # Wrap in a styled div with only Arial font
            styled_html = f"""
            <style>
                * {{
                    font-family: Arial !important;
                }}
            </style>
            <div style="font-family: Arial; padding: 20px; line-height: 1.6;">
            {html}
            </div>
            """
            return styled_html
        except Exception as e:
            return f"<p>Error loading documentation: {str(e)}</p>"
    
    def setup_ui(self):
        """Setup the manual dialog UI"""
        layout = QVBoxLayout()
        
        # Create tab widget for different manual sections
        tab_widget = QTabWidget()
        
        # Overview Tab
        overview_tab = self.create_overview_tab()
        tab_widget.addTab(overview_tab, "Overview")

        # Plotting Tab
        plotting_tab = self.create_plotting_tab()
        tab_widget.addTab(plotting_tab, "Plotting")
        
        # CPV Mode Tab
        cpv_tab = self.create_cpv_tab()
        tab_widget.addTab(cpv_tab, "CPV Mode")
        
        # DRC Mode Tab
        drc_tab = self.create_drc_tab()
        tab_widget.addTab(drc_tab, "DRC Mode")

        # Loop-defined Mode Tab
        loopdefined_tab = self.create_loopdefined_tab()
        tab_widget.addTab(loopdefined_tab, "Loop-defined Mode")

        # Drag and drop Mode Tab
        dragdrop_tab = self.create_dragdrop_tab()
        tab_widget.addTab(dragdrop_tab, "Drag and Drop Mode")

        # Help Tab
        help_tab = self.create_help_tab()
        tab_widget.addTab(help_tab, "Help")
        
        layout.addWidget(tab_widget)
        
        # Close button
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def create_overview_tab(self):
        """Create overview tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_overview.md")
        browser.setHtml(html_content)
        
        layout.addWidget(browser)
        widget.setLayout(layout)
        return widget

    def create_plotting_tab(self):
        """Create plotting tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_plotting.md")
        browser.setHtml(html_content)
        
        layout.addWidget(browser)
        widget.setLayout(layout)
        return widget
    
    def create_cpv_tab(self):
        """Create CPV mode manual tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_cpv.md")
        browser.setHtml(html_content)
        
        layout.addWidget(browser)
        widget.setLayout(layout)
        return widget
    
    def create_drc_tab(self):
        """Create DRC mode manual tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_drc.md")
        browser.setHtml(html_content)
        
        layout.addWidget(browser)
        widget.setLayout(layout)
        return widget
    
    def create_loopdefined_tab(self):
        """Create Loop-defined mode manual tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_loopdefined.md")
        browser.setHtml(html_content)
        
        layout.addWidget(browser)
        widget.setLayout(layout)
        return widget

    def create_dragdrop_tab(self):
        """Create Drag and Drop mode manual tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_dragdrop.md")
        browser.setHtml(html_content)
        
        layout.addWidget(browser)
        widget.setLayout(layout)
        return widget

    def create_help_tab(self):
        """Create help tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_help.md")
        browser.setHtml(html_content)

        layout.addWidget(browser)
        widget.setLayout(layout)
        return widget