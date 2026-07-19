import os
import json


class SessionState:

    def __init__(self, workdir: str):
        self.workdir = workdir
        self.open_files: dict = {}
        self.history: list = []
        self.last_action: str = "None"

    def register_file(self, filename: str, full_path: str, note: str):
        extension = os.path.splitext(filename)[1].lower()
        self.open_files[filename] = {
            "extension": extension,
            "full_path": full_path,
            "note": note
        }

    def set_last_action(self, action: str):
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

    def remove_file(self, filename: str):  
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
        history = data.get("history", [])
        history = self._sanitize_history(history)
        self.history = history

    def _sanitize_history(self, history: list) -> list:

            while history and history[0]["role"] != "user":
                history.pop(0)

            clean = []
            i = 0
            while i < len(history):
                msg = history[i]
                if msg["role"] == "assistant" and msg.get("tool_calls"):

                    if i + 1 < len(history) and history[i + 1]["role"] == "tool":
                        clean.append(msg)
                    else:
                        i += 1
                        continue
                else:
                    clean.append(msg)
                i += 1

            return clean
