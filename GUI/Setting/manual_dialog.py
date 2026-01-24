"""
Manual Dialog Module
Provides user manual and documentation
"""
import os
import sys
import re
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
    
    def load_markdown_file(self, filename):
        """Load and convert markdown file to HTML, removing all font-family declarations"""
        filepath = os.path.join(self.docs_dir, filename)
        
        if not os.path.exists(filepath):
            return f"<p>Error: Documentation file not found: {filepath}</p>"
        
        try:
            # Lazy import markdown to avoid py2app zipimporter issues
            import markdown
            
            with open(filepath, 'r', encoding='utf-8') as f:
                markdown_content = f.read()
            
            # Convert markdown to HTML
            # Note: 'codehilite' extension requires pygments which can cause issues in py2app
            # Using 'extra' and 'nl2br' for better compatibility and line breaks
            # 'nl2br' converts single line breaks to <br> tags
            html = markdown.markdown(markdown_content, extensions=['extra', 'nl2br'])
            
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

        # User-defined Mode Tab
        usrdefined_tab = self.create_usrdefined_tab()
        tab_widget.addTab(usrdefined_tab, "User-defined Mode")
        
        # Troubleshooting Tab
        trouble_tab = self.create_troubleshooting_tab()
        tab_widget.addTab(trouble_tab, "Troubleshooting")

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
    
    def create_usrdefined_tab(self):
        """Create User-defined mode manual tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_usrdefined.md")
        browser.setHtml(html_content)
        
        layout.addWidget(browser)
        widget.setLayout(layout)
        return widget
    
    def create_troubleshooting_tab(self):
        """Create troubleshooting tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setFontFamily("Arial")
        
        # Load markdown file and convert to HTML
        html_content = self.load_markdown_file("manual_troubleshooting.md")
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