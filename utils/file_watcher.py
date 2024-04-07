import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path


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


def watch_file(folder_path: str, file_name: str, input_function):
    file_path = Path(folder_path) / file_name
    event_handler = FileWatcher(input_function, file_path)
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
    folder_path = r"C:\Users\Fetyani\Desktop\Maher"
    watch_file(folder_path, "test.xlsx", your_input_function)
