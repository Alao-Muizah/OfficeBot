import os 
import json 

class SessionState:

    def __init__(self, workdir:str):
        self.workdir = workdir
        self.open_files: dict = {}
        self.history: list = []
        self.last_action: str = "None"

    
    def register_file(self, filename:str, full_path:str, note:str):

        extension = os.path.splitext(filename)[1].lower()

        self.open_files[filename] = {
            "extension": extension,
            "full_path": full_path,
            "note": note
        }

    def set_last_action(self, action:str):

        self.last_action = action 
    
    def get_context_summary(self) -> str:
        files_summary = "  No files worked with yet."

        if self.open_files:
            lines = []
            for name, info in self.open_files.items():
                lines.append(f"  - {name} ({info['extension']}) — {info['note']}")
            files_summary = "\n".join(lines)

        return (
            f"Files worked with this session:\n"
            f"{files_summary}\n"
            f"Last action: {self.last_action}"
        )
    def remove_file(self, filename:str):

        self.open_files.pop(filename, None)

    def clear(self): 
        self.open_files = {}
        self.history = []
        self.last_action = "None"


    def save_to_file(self, filepath: str):
        
        data = {
            "open_files": self.open_files,
            "last_action": self.last_action,
            "history": self.history
        }
        with open(filepath, "w") as f:
            json.dump(data, f, indent=2)

    def load_from_file(self, filepath: str):
        
        if not os.path.exists(filepath):
            return  
        with open(filepath) as f:
            data = json.load(f)
        self.open_files = data.get("open_files", {})
        self.last_action = data.get("last_action", "None")
        self.history = data.get("history", [])