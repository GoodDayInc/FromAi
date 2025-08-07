import customtkinter as ctk

class PlaceholderEntry(ctk.CTkEntry):
    """An entry widget that shows placeholder text."""

    def __init__(self, master=None, placeholder="", **kwargs):
        super().__init__(master, **kwargs)
        self.placeholder = placeholder
        self.bind("<FocusIn>", self._clear_placeholder)
        self.bind("<FocusOut>", self._add_placeholder)
        self._add_placeholder()

    def _clear_placeholder(self, e):
        if self.get() == self.placeholder:
            self.delete(0, "end")

    def _add_placeholder(self, e=None):
        if not self.get():
            self.insert(0, self.placeholder)

    def get(self) -> str:
        """Get the entry text, returning empty string if it's the placeholder."""
        text = super().get()
        if text == self.placeholder:
            return ""
        return text


class Tooltip:
    """A modern, theme-aware tooltip for customtkinter widgets."""

    def __init__(self, widget, text: str, delay: int = 500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tooltip_window = None
        self.id = None
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
        self.widget.bind("<ButtonPress>", self.on_leave)

    def on_enter(self, event=None):
        self.schedule_tooltip()

    def on_leave(self, event=None):
        self.hide_tooltip()

    def schedule_tooltip(self):
        self.hide_tooltip()
        self.id = self.widget.after(self.delay, self.show_tooltip)

    def show_tooltip(self):
        if self.tooltip_window or not self.text:
            return

        # Get widget's theme colors
        bg_color = self.widget.cget("fg_color")
        if isinstance(bg_color, (list, tuple)):
             bg_color = bg_color[1] # Use the hover color or second color

        fg_color = self.widget.cget("text_color")

        self.tooltip_window = ctk.CTkToplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)

        label = ctk.CTkLabel(
            self.tooltip_window,
            text=self.text,
            corner_radius=5,
            fg_color=("#333333", "#CCCCCC"),
            text_color=("#FFFFFF", "#000000"),
        )
        label.pack(padx=1, pady=1)

        self.tooltip_window.update_idletasks()

        # Position the tooltip
        widget_x = self.widget.winfo_rootx()
        widget_y = self.widget.winfo_rooty()
        widget_height = self.widget.winfo_height()
        tip_width = self.tooltip_window.winfo_width()
        tip_height = self.tooltip_window.winfo_height()

        x = widget_x + self.widget.winfo_width() // 2 - tip_width // 2
        y = widget_y + widget_height + 5

        # Adjust position to stay on screen
        screen_width = self.widget.winfo_screenwidth()
        screen_height = self.widget.winfo_screenheight()
        if x + tip_width > screen_width:
            x = screen_width - tip_width - 10
        if x < 0:
            x = 10
        if y + tip_height > screen_height:
            y = widget_y - tip_height - 10
        if y < 0:
            y = 10

        self.tooltip_window.wm_geometry(f"+{int(x)}+{int(y)}")

    def hide_tooltip(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
