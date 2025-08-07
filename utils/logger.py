import datetime
import customtkinter

class Logger:
    """Manages logging to the GUI's text widget."""

    def __init__(self, output_widget: customtkinter.CTkTextbox):
        self.output_widget = output_widget

    def log(self, message: str, level: str = "info") -> None:
        if self.output_widget is None:
            return

        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"

        # The CTkTextbox needs to be configured to be editable before inserting text
        self.output_widget.configure(state="normal")
        self.output_widget.insert("end", formatted_message, level)
        self.output_widget.configure(state="disabled")
        self.output_widget.see("end")

    def info(self, message: str) -> None:
        self.log(message, "info")

    def success(self, message: str) -> None:
        self.log(f"✓ {message}", "success")

    def warning(self, message: str) -> None:
        self.log(f"⚠ {message}", "warning")

    def error(self, message: str) -> None:
        self.log(f"✗ {message}", "error")
