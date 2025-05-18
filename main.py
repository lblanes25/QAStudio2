#!/usr/bin/env python3
"""
Main entry point for the QA Analytics Automation Framework UI.
"""

import tkinter as tk
from qa_analytics.enhanced_qa_analytics_app import QAAnalyticsApp

def main():
    """Main entry point for the application"""
    root = tk.Tk()
    app = QAAnalyticsApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()