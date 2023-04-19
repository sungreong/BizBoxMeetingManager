from PyQt5.QtWidgets import (
    QMessageBox,
)
# TODO:
def handle_error(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            QMessageBox(f"An error occurred: {str(e)}")
    return wrapper