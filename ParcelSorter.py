import threading
import queue
import csv
import os
from Setup import *
from Operation import *
import winsound
alarm_list = ["207: PARCEL_LOST", "203: DELIVERED", "NO_STATUS", "217: TRANSSHIPMENT_COMPLETE"]

CSV_FILE = "scans.csv"

parcel_queue = queue.Queue()
scan_index = 0
scan_lock = threading.Lock()
csv_lock = threading.Lock()

CSV_HEADERS = [
    "scan_index",
    "parcel_id",
    "sub_batch",
    "status",
    "driver_id",
    "warehouse",
    "segment",
    "storage",
    "driver_memo"
]

def write_csv(scan_index, data):
    with open(CSV_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        writer.writerow({
            "scan_index": scan_index,
            "parcel_id": data["parcel_id"],
            "sub_batch": data["sub_batch"],
            "status": data["status"],
            "driver_id": data["driver_id"],
            "warehouse": data["warehouse"],
            "segment": data["segment"],
            "storage": data["storage"],
            "driver_memo": data["driver_memo"],
        })

def init_csv():
    if not os.path.exists(CSV_FILE):
        with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
            writer.writeheader()

def parcel_worker(driver):
    global scan_index
    while True:
        parcel_id = parcel_queue.get()

        if parcel_id is None:
            parcel_queue.task_done()
            break

        with scan_lock:
            scan_index += 1
            current_index = scan_index

        #print(f"🔄 Processing #{current_index}: {parcel_id}")

        try:
            data = search_parcel(driver, parcel_id)

            with csv_lock:
                write_csv(current_index, data)
            print(
                f"✅ #{current_index} | "
                f"{data['parcel_id']} | "
                f"{data['sub_batch']} | "
                f"{data['status']} | "
                f"{data['driver_id']} | "
                f"{data['warehouse']} | "
                f"{data['segment']} | "
                f"{data['storage']} | "
                f"{data['driver_memo']}"
            )
            if data['status'] in alarm_list:
                winsound.Beep(1000, 1000)   # 1000Hz for 1 second
            if data['warehouse'] != 'BUF Warehouse':
                winsound.Beep(2000, 1000)   # 1000Hz for 1 second

            # print(
            #     f"✅ #{current_index} | "
            #     f"{data['parcel_id']} | "
            #     f"{data['sub_batch']} | "
            #     f"{data['status']} | "
            # )

        except Exception as e:
            print(f"❌ Error processing {parcel_id}: {e}")

        parcel_queue.task_done()



if __name__ == "__main__":
    driver = init_edge_driver()
    open_new_tab(driver)
    open_edit_order(driver)
    init_csv()

    worker = threading.Thread(
        target=parcel_worker,
        args=(driver,),
    )
    worker.start()


    print("📦 Parcel Sorting System Ready")
    print("Scan parcels continuously (Ctrl+C to stop)\n")

    try:
        while True:
            parcel_id = input().strip()
            if parcel_id:
                parcel_queue.put(parcel_id)
                print(f"📥 Queued: {parcel_id}")

    except KeyboardInterrupt:
        print("\n🛑 Shutting down...")

        if worker.is_alive():
            parcel_queue.put(None)
            worker.join()

        driver.quit()