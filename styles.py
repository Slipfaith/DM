MAIN_STYLE = """
QFrame#DropArea {
    border: 2px dashed #cccccc;
    border-radius: 10px;
    padding: 20px;
    background-color: #f9f9f9;
}

#dragDropArea {
    background-color: #ffffff;
    border: 2px dashed #cccccc;
    border-radius: 8px;
    min-height: 200px;
    margin: 20px;
}

#dragDropArea:hover {
    border-color: #999999;
    background-color: #fafafa;
}

#dragDropArea[dragActive="true"] {
    border-color: #4CAF50;
    background-color: #f1f8f4;
}

#dragDropText {
    color: #666666;
    font-size: 16px;
    font-weight: 500;
}

#contentWidget {
    background-color: #ffffff;
}

#fileGroup {
    font-weight: bold;
    border: 1px solid #e0e0e0;
    border-radius: 4px;
    padding: 10px;
}

#fileList {
    border: 1px solid #e0e0e0;
    border-radius: 4px;
    background-color: #fafafa;
    padding: 5px;
    min-height: 150px;
}

#fileList::item {
    padding: 5px;
    border-bottom: 1px solid #f0f0f0;
}

#fileList::item:hover {
    background-color: #f5f5f5;
}

#fileList::item:selected {
    background-color: #e3f2fd;
    color: #1976d2;
}

#statusLabel {
    color: #666666;
    font-size: 14px;
    padding: 10px;
    background-color: #f5f5f5;
    border-radius: 4px;
}

#progressBar {
    border: none;
    border-radius: 4px;
    background-color: #e0e0e0;
    text-align: center;
    height: 24px;
}

#progressBar::chunk {
    background-color: #4CAF50;
    border-radius: 4px;
}

#processButton {
    background-color: #4CAF50;
    color: white;
    border: none;
    padding: 10px 24px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 14px;
    min-width: 120px;
}

#processButton:hover {
    background-color: #45a049;
}

#processButton:pressed {
    background-color: #3d8b40;
}

#processButton:disabled {
    background-color: #cccccc;
    color: #666666;
}

#clearButton {
    background-color: #f44336;
    color: white;
    border: none;
    padding: 10px 24px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 14px;
    min-width: 120px;
}

#clearButton:hover {
    background-color: #da190b;
}

#clearButton:pressed {
    background-color: #ba000d;
}

#clearButton:disabled {
    background-color: #cccccc;
    color: #666666;
}

#colorInput {
    padding: 5px;
    border: 1px solid #cccccc;
    border-radius: 4px;
    background-color: white;
    min-width: 100px;
}

#colorInput:focus {
    border-color: #4CAF50;
}

#dryRunCheck {
    color: #666666;
    font-size: 14px;
    padding: 5px;
}

#dryRunCheck::indicator {
    width: 18px;
    height: 18px;
}

#dryRunCheck::indicator:unchecked {
    border: 2px solid #cccccc;
    border-radius: 3px;
    background-color: white;
}

#dryRunCheck::indicator:checked {
    border: 2px solid #4CAF50;
    border-radius: 3px;
    background-color: #4CAF50;
    image: url(check.png);
}

#logText {
    border: 1px solid #e0e0e0;
    border-radius: 4px;
    background-color: #f8f8f8;
    font-family: 'Consolas', 'Monaco', monospace;
    font-size: 12px;
    padding: 10px;
    min-height: 150px;
}

QLabel {
    color: #333333;
}

QGroupBox {
    font-size: 14px;
    color: #333333;
}

#summaryLabel {
    color: #333333;
    font-size: 13px;
    padding: 10px;
    background-color: #e8f5e9;
    border: 1px solid #4CAF50;
    border-radius: 4px;
}

#summaryLabel a {
    color: #1976d2;
    text-decoration: underline;
}

#progressLabel {
    color: #666666;
    font-size: 12px;
}
"""
