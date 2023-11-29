import watchdog.events
import watchdog.observers
import time
from threading import Thread
import threading
from commons.config import file_operator_classes
import xlrd
import xlwings as xw

class Handler(watchdog.events.PatternMatchingEventHandler):
    def __init__(self, file_operator_class):
        self.file_operator_class = file_operator_class
        input_file_pattern_list = file_operator_class.input_file_pattern_list
        print(input_file_pattern_list)
        watchdog.events.PatternMatchingEventHandler.__init__(self, patterns=input_file_pattern_list,
                                                             ignore_directories=True, case_sensitive=False)
                          
    def on_created(self, event): 
        file_operator = self.file_operator_class(event.src_path)
        process_file_thread = Thread(target=file_operator.process_file)                                    
        process_file_thread.start()   

    def on_deleted(self, event):
        print("Deleted File: %s" % event.src_path)

if __name__ == "__main__":
    # app = xw.App(visible=False)

    report_type = "PBS"
    file_operator_class = file_operator_classes[report_type]
    src_path = file_operator_class.base_folder

    print(f"Report Type: {report_type}")
    print(f"Source Path: {src_path}")

    event_handler = Handler(file_operator_class)
    observer = watchdog.observers.Observer()
    observer.schedule(event_handler, path=src_path, recursive=True)
    observer.start()
    
    try:
        while True:
            print(f"Active Operator Threads: {threading.active_count() - 3}")
            if threading.active_count() == 3:
                print(f"Open Excel: {len(xw.apps)}")
                xl_keys = xw.apps.keys()
                if len(xw.apps) > 0:
                    for xl_key in xl_keys:
                        xw.apps[xl_key].quit()
            time.sleep(5)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()