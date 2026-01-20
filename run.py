#!/usr/bin/env python3
"""
Run the KCSE Enterprise System Web Application
"""

from app import app

if __name__ == '__main__':
    # Production settings
    app.run(
        host='0.0.0.0',
        port=5000,
        debug=False
    )
