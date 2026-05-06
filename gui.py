import customtkinter as ctk
import threading
import assistant

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AssistantGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Jarvis")
        self.geometry("600x700")
        
        # Grid Layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Chat History
        self.chat_history = ctk.CTkTextbox(self, state="disabled", wrap="word", font=("Inter", 14))
        self.chat_history.grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 10), sticky="nsew")
        
        # Status Label
        self.status_label = ctk.CTkLabel(self, text="Jarvis is online.", text_color="gray", font=("Inter", 12))
        self.status_label.grid(row=1, column=0, columnspan=2, padx=20, pady=0, sticky="w")
        
        # Input Frame
        self.input_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.input_frame.grid(row=2, column=0, columnspan=2, padx=20, pady=(10, 20), sticky="ew")
        self.input_frame.grid_columnconfigure(0, weight=1)
        
        self.entry = ctk.CTkEntry(self.input_frame, placeholder_text="Type your command...", font=("Inter", 14))
        self.entry.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.entry.bind("<Return>", lambda event: self.send_text_command())
        
        self.send_button = ctk.CTkButton(self.input_frame, text="Send", width=70, command=self.send_text_command)
        self.send_button.grid(row=0, column=1, padx=(0, 10))
        
        self.mic_button = ctk.CTkButton(self.input_frame, text="🎙️", width=40, command=self.start_listening, fg_color="#E07A5F", hover_color="#C06A4F")
        self.mic_button.grid(row=0, column=2, padx=(0, 10))
        
        self.interrupt_button = ctk.CTkButton(self.input_frame, text="🛑", width=40, command=self.interrupt_speech, fg_color="#E63946", hover_color="#D62828")
        self.interrupt_button.grid(row=0, column=3)
        
        self.interrupt_event = threading.Event()
        self.is_processing = False
        
        self.append_to_chat("Jarvis: Jarvis is now online. How can I help you today?")
        
    def append_to_chat(self, text):
        self.chat_history.configure(state="normal")
        self.chat_history.insert("end", text + "\n\n")
        self.chat_history.configure(state="disabled")
        self.chat_history.yview("end")
        
    def set_status(self, text):
        self.status_label.configure(text=text)

    def send_text_command(self):
        if self.is_processing:
            return
        command = self.entry.get().strip()
        if not command:
            return
            
        self.entry.delete(0, "end")
        self.append_to_chat(f"You: {command}")
        
        self.is_processing = True
        self.set_status("Thinking...")
        
        # Process command in a background thread
        threading.Thread(target=self.process_and_speak, args=(command,), daemon=True).start()
        
    def start_listening(self):
        if self.is_processing:
            return
            
        self.is_processing = True
        self.set_status("Listening...")
        
        # Listen in a background thread
        threading.Thread(target=self.listen_and_process, daemon=True).start()
        
    def listen_and_process(self):
        def update_status(msg):
            self.after(0, self.set_status, msg)
            
        command = assistant.listen(status_callback=update_status)
        
        if command:
            self.after(0, self.append_to_chat, f"You: {command}")
            self.after(0, self.set_status, "Thinking...")
            self.process_and_speak(command)
        else:
            # Delay resetting to "Ready." so the user can read any error messages from listen()
            def reset_ready():
                if self.status_label.cget("text") not in ["Thinking...", "Speaking...", "Listening..."]:
                    self.set_status("Ready.")
            self.after(3000, reset_ready)
            self.is_processing = False
            
    def process_and_speak(self, command):
        response = assistant.process_command(command)
        
        if response:
            self.after(0, self.append_to_chat, f"Jarvis: {response}")
            self.after(0, self.set_status, "Speaking...")
            
            self.interrupt_event.clear()
            assistant.speak(response, self.interrupt_event)
            
        self.after(0, self.set_status, "Ready.")
        self.is_processing = False
        
    def interrupt_speech(self):
        self.interrupt_event.set()
        self.set_status("Speech interrupted.")

if __name__ == "__main__":
    app = AssistantGUI()
    app.mainloop()
