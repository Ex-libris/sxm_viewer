"""Shared Qt style strings for the GUI."""

def lower_control_frame_style(border: str, bg: str) -> str:
    return f"QFrame#lowerControlFrame {{ border-top: 1px solid {border}; background-color: {bg}; }}"


def mode_selector_style(mode_border: str, mode_text: str, mode_checked: str,
                        mode_checked_text: str = "#ffffff") -> str:
    return (
        "QWidget#modeSelector QToolButton {"
        f" border: 1px solid {mode_border};"
        " padding: 6px 12px;"
        " background: transparent;"
        f" color: {mode_text};"
        "}"
        "QWidget#modeSelector QToolButton:checked {"
        f" background: {mode_checked};"
        f" color: {mode_checked_text};"
        "}"
        "QWidget#modeSelector QToolButton + QToolButton {"
        " border-left: none;"
        "}"
    )


MAIN_SHORTCUTS_PANEL_STYLE = """
QFrame#shortcutsPanel {
    background: rgba(64, 96, 160, 25%);
    border: 1px solid rgba(64, 96, 160, 60%);
    border-radius: 8px;
    padding: 6px;
}
"""

MAIN_TOOLBAR_CANVAS_BUTTON_STYLE = (
    "QPushButton {"
    " background-color: #2563eb;"
    " color: #ffffff;"
    " font-weight: 600;"
    " padding: 4px 10px;"
    " border-radius: 6px;"
    "}"
    "QPushButton:hover { background-color: #1d4ed8; }"
    "QPushButton:pressed { background-color: #1e40af; }"
)

CANVAS_DIALOG_STYLE = """
QDialog {
    background-color: #2b2b2b;
    color: #f0f0f0;
}
QGroupBox {
    border: 1px solid #444444;
    border-radius: 6px;
    margin-top: 10px;
}
QGroupBox:title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
    color: #e6e6e6;
}
QLabel {
    color: #e6e6e6;
}
QLineEdit, QComboBox, QSpinBox {
    background-color: #1f1f1f;
    border: 1px solid #444444;
    padding: 4px;
    border-radius: 4px;
    color: #f0f0f0;
}
QPushButton {
    background-color: #3a3a3a;
    border: 1px solid #555555;
    border-radius: 4px;
    padding: 4px 8px;
    color: #f0f0f0;
}
QPushButton:hover {
    background-color: #444444;
}
QPushButton:pressed {
    background-color: #2f2f2f;
}
QCheckBox {
    color: #e6e6e6;
}
QScrollArea {
    background: transparent;
}
"""

# Light theme variant for canvas dialogs
CANVAS_DIALOG_STYLE_LIGHT = """
QDialog {
    background-color: #f7f7f7;
    color: #1b1b1b;
}
QGroupBox {
    border: 1px solid #c4c4c4;
    border-radius: 6px;
    margin-top: 10px;
}
QGroupBox:title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
    color: #303030;
}
QLabel {
    color: #303030;
}
QLineEdit, QComboBox, QSpinBox {
    background-color: #ffffff;
    border: 1px solid #c4c4c4;
    padding: 4px;
    border-radius: 4px;
    color: #1b1b1b;
}
QPushButton {
    background-color: #e6e6e6;
    border: 1px solid #c4c4c4;
    border-radius: 4px;
    padding: 4px 8px;
    color: #1b1b1b;
}
QPushButton:hover {
    background-color: #dcdcdc;
}
QPushButton:pressed {
    background-color: #cfcfcf;
}
QCheckBox {
    color: #303030;
}
QScrollArea {
    background: transparent;
}
"""

