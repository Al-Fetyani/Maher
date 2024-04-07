import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path
import threading


class FileWatcher(FileSystemEventHandler):
    def __init__(self, input_function, file_path: Path):
        self.input_function = input_function
        self.file_path = file_path
        self.is_modified = False

    def on_modified(self, event):
        if not event.is_directory and Path(event.src_path) == self.file_path:
            if not self.is_modified:
                self.is_modified = True
                self.input_function()
            elif self.is_modified:
                self.is_modified = False


def start_watching(file_path: str, input_function):
    thread = threading.Thread(
        target=watch_file, args=(file_path, input_function), daemon=True
    )
    thread.start()
    return thread


def watch_file(file_path: str, input_function):
    folder_path = Path(file_path).parent
    event_handler = FileWatcher(input_function, Path(file_path))
    observer = Observer()
    observer.schedule(event_handler, path=folder_path, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(300)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


i = 1


def your_input_function():
    global i
    print(f"File modified {i} times")
    i += 1


if __name__ == "__main__":
    file_path = r"C:\Users\Fetyani\Desktop\Maher\test.xlsx"
    watch_file(file_path, your_input_function)
