import tkinter as tk
from tkinter import filedialog
import time
import threading
import docx
import keyboard

class AutoTypingApp:
    def __init__(self, root):
        """
        Initialize the Auto Typing Application.

        Parameters:
        - root: The root window of the GUI.
        """
        self.root = root
        self.root.title("Auto Typing App")
        self.root.geometry("400x400")
        self.root.configure(bg="black")  # Set background color

        # Create and place GUI elements
        self.heading_label = tk.Label(self.root, text="Auto Typing App", font=("Courier", 18, "bold"), fg="#00ff00", bg="black")
        self.heading_label.pack(pady=10)

        self.text_box = tk.Text(self.root, height=10, width=40, font=("Courier", 12), bg="black", fg="#00ff00")
        self.text_box.pack(pady=10)

        self.upload_button = tk.Button(self.root, text="Upload File", command=self.upload_file, font=("Courier", 12), bg="black", fg="black")
        self.upload_button.pack(pady=5)

        self.start_button = tk.Button(self.root, text="Start Typing", command=self.start_typing, font=("Courier", 12), bg="black", fg="black")
        self.start_button.pack(pady=5)

        self.pause_button = tk.Button(self.root, text="Pause Typing", command=self.pause_typing, font=("Courier", 12), bg="black", fg="black")
        self.pause_button.pack(pady=5)

        # Initialize instance variables
        self.typing_thread = None
        self.typing_paused = False
        self.countdown_label = None

    def upload_file(self):
        """
        Upload a Word document and display its content in the text box.
        """
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path:
            doc = docx.Document(file_path)
            content = "\n".join([para.text for para in doc.paragraphs])
            self.text_box.delete(1.0, tk.END)
            self.text_box.insert(tk.END, content)

    def start_typing(self):
        """
        Start the typing process with a countdown and initiate a typing thread.
        """
        self.countdown_label = tk.Label(self.root, font=("Helvetica", 14), fg="#ecf0f1", bg="#2c3e50")
        self.countdown_label.pack(pady=10)
        self.typing_thread = threading.Thread(target=self.typing_process)
        self.typing_thread.start()

    def typing_process(self):
        """
        Simulate the typing process with a pause and resume functionality.
        """
        for i in range(4, 0, -1):
            if self.typing_paused:
                self.countdown_label.config(text="Paused")
                while self.typing_paused:
                    time.sleep(1)
            else:
                self.countdown_label.config(text=f"Starting in {i} seconds")
                time.sleep(1)

        self.countdown_label.config(text="Typing...")

        text_to_type = self.text_box.get(1.0, tk.END)
        lines = text_to_type.split("\n")
        for line in lines:
            if not self.typing_paused:
                keyboard.write(line)
                keyboard.press_and_release("enter")
                time.sleep(0.5)
            else:
                break

        self.countdown_label.config(text="Typing finished")

    def pause_typing(self):
        """
        Toggle the pause and resume of the typing process.
        """
        self.typing_paused = not self.typing_paused
        if self.typing_paused:
            self.pause_button.config(text="Resume Typing")
        else:
            self.pause_button.config(text="Pause Typing")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoTypingApp(root)
    root.mainloop()
