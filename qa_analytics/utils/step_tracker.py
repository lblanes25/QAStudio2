import tkinter as tk
from tkinter import ttk


class StepTracker(ttk.Frame):
    """
    Visual step tracker component for wizards and multi-step processes.

    Displays a horizontal series of numbered steps with connecting lines,
    and visually indicates which steps are completed, current, or upcoming.
    """

    def __init__(self, parent, steps, initial_step=0):
        """
        Initialize the step tracker.

        Args:
            parent: Parent tkinter container
            steps: List of step names to display
            initial_step: Index of the initial active step (0-based)
        """
        super().__init__(parent)
        self.steps = steps
        self.current_step = initial_step

        # Colors for different states
        self.colors = {
            'completed': {
                'fill': '#90EE90',  # Light green
                'outline': '#228B22',  # Forest green
                'text': 'black'
            },
            'current': {
                'fill': '#CCE5FF',  # Light blue
                'outline': '#0066CC',  # Medium blue
                'text': 'black'
            },
            'future': {
                'fill': 'white',
                'outline': 'gray',
                'text': 'gray'
            }
        }

        self._setup_ui()

    def _setup_ui(self):
        """Create the step tracker UI"""
        # Step indicators and connectors
        self.indicators = []

        for i, step_name in enumerate(self.steps):
            # Create step container
            step_frame = ttk.Frame(self)

            # For all except the last step, add connector line
            if i < len(self.steps) - 1:
                step_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)

                # Create connector between steps (will be added after indicator)
                connector_frame = ttk.Frame(step_frame)
                connector_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)

                connector = ttk.Separator(connector_frame, orient="horizontal")
                connector.pack(fill=tk.X, expand=True, pady=15)
            else:
                # Last step doesn't need a connector
                step_frame.pack(side=tk.LEFT, padx=(0, 10))

            # Create circular indicator
            indicator_frame = ttk.Frame(step_frame)
            indicator_frame.pack(side=tk.LEFT, padx=(10, 5))

            # Use Canvas for custom drawing
            # FIX: Don't try to get bg color from master, use system default instead
            indicator = tk.Canvas(indicator_frame, width=30, height=30,
                                  highlightthickness=0)
            indicator.pack()

            # Draw circle and number
            circle = indicator.create_oval(5, 5, 25, 25, width=2)
            text = indicator.create_text(15, 15, text=str(i + 1))

            # Add step label
            label = ttk.Label(step_frame, text=step_name)
            label.pack(side=tk.LEFT, padx=(0, 5))

            # Store references for updating
            self.indicators.append((indicator, circle, text, label))

        # Update initial state
        self.set_current_step(self.current_step)

    def set_current_step(self, step_index):
        """
        Update the visual state of the step tracker.

        Args:
            step_index: Index of the current step (0-based)
        """
        if step_index < 0 or step_index >= len(self.steps):
            return  # Invalid step index

        self.current_step = step_index

        for i, (indicator, circle, text, label) in enumerate(self.indicators):
            if i < step_index:
                # Completed step
                state = 'completed'
                label_font = ('TkDefaultFont', 9, 'normal')
            elif i == step_index:
                # Current step
                state = 'current'
                label_font = ('TkDefaultFont', 9, 'bold')
            else:
                # Future step
                state = 'future'
                label_font = ('TkDefaultFont', 9, 'normal')

            # Update colors based on state
            indicator.itemconfig(
                circle,
                fill=self.colors[state]['fill'],
                outline=self.colors[state]['outline']
            )
            indicator.itemconfig(
                text,
                fill=self.colors[state]['text']
            )

            # Update label style
            label.configure(
                font=label_font,
                foreground=self.colors[state]['text']
            )

    def get_current_step(self):
        """
        Get the current step index.

        Returns:
            int: Current step index (0-based)
        """
        return self.current_step

    def next_step(self):
        """
        Move to the next step if possible.

        Returns:
            bool: True if moved to next step, False if already at last step
        """
        if self.current_step < len(self.steps) - 1:
            self.set_current_step(self.current_step + 1)
            return True
        return False

    def previous_step(self):
        """
        Move to the previous step if possible.

        Returns:
            bool: True if moved to previous step, False if already at first step
        """
        if self.current_step > 0:
            self.set_current_step(self.current_step - 1)
            return True
        return False


# Example usage (for testing)
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Step Tracker Example")
    root.geometry("800x200")

    frame = ttk.Frame(root, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # Create step tracker with 4 steps
    steps = ["Select Template", "Basic Configuration", "Set Parameters", "Review & Save"]
    tracker = StepTracker(frame, steps)
    tracker.pack(fill=tk.X, pady=20)

    # Add controls for testing
    control_frame = ttk.Frame(frame)
    control_frame.pack(pady=20)

    ttk.Button(
        control_frame,
        text="Previous",
        command=tracker.previous_step
    ).pack(side=tk.LEFT, padx=5)

    ttk.Button(
        control_frame,
        text="Next",
        command=tracker.next_step
    ).pack(side=tk.LEFT, padx=5)

    root.mainloop()