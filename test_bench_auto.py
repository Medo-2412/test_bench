import tkinter as tk
from tkinter import messagebox, ttk
import threading
import time
import serial
import serial.tools.list_ports
from openpyxl import Workbook
from datetime import datetime

# Main Window
window = tk.Tk()
window.title("Test Bench")
window.geometry("550x800")

# ====== Validators ======

# Validates all numeric fields with max=100
def is_valid_integer_input(value):
    if value == "":
        return True  # allow temporary blank while typing
    return value.isdigit() and int(value) <= 100

# Validates that ending point > starting point and ≤ 100
def validate_end_input(end_value):
    if end_value == "":
        return True
    if not end_value.isdigit():
        return False
    end = int(end_value)
    # Allow partial inputs like '3' when typing '30'
    if len(end_value) < 2:
        return True
    try:
        start = int(start_entry.get())
    except:
        start = 0
    return end > start and end <= 100


vcmd = (window.register(is_valid_integer_input), "%P")
end_vcmd = (window.register(validate_end_input), "%P")

# ====== Serial Ports ======
def list_serial_ports():
    ports = serial.tools.list_ports.comports()
    return [port.device for port in ports]

tk.Label(window, text="Select Port:").pack()
port_combo = ttk.Combobox(window, values=list_serial_ports(), width=37, state="readonly")
port_combo.pack()

def refresh_ports():
    ports = list_serial_ports()
    port_combo['values'] = ports
    if ports:
        port_combo.current(0)

# ====== Inputs ======
tk.Button(window, text="Refresh Ports", command=refresh_ports).pack(pady=5)

tk.Label(window, text="Starting Point (0–100):").pack()
start_entry = tk.Entry(window, width=40, validate="key", validatecommand=vcmd)
start_entry.insert(0, "0")
start_entry.pack()

tk.Label(window, text="Ending Point (must be > Starting Point, max 100):").pack()
end_entry = tk.Entry(window, width=40, validate="key", validatecommand=end_vcmd)
end_entry.insert(0, "0")
end_entry.pack()

tk.Label(window, text="Increment Amount:").pack()
increment_entry = tk.Entry(window, width=40, validate="key", validatecommand=vcmd)
increment_entry.insert(0, "0")
increment_entry.pack()

tk.Label(window, text="Time Delay (seconds):").pack()
delay_entry = tk.Entry(window, width=40, validate="key", validatecommand=vcmd)
delay_entry.insert(0, "0")
delay_entry.pack()

tk.Label(window, text="Final Hold Time (minutes):").pack()
final_hold_entry = tk.Entry(window, width=40, validate="key", validatecommand=vcmd)
final_hold_entry.insert(0, "0")
final_hold_entry.pack()

tk.Label(window, text="Ramp-Down Delay (seconds):").pack()
ramp_down_delay_entry = tk.Entry(window, width=40, validate="key", validatecommand=vcmd)
ramp_down_delay_entry.insert(0, "0")
ramp_down_delay_entry.pack()

tk.Label(window, text="Ramp-Down Decrement (positive integer):").pack()
ramp_down_decrement_entry = tk.Entry(window, width=40, validate="key", validatecommand=vcmd)
ramp_down_decrement_entry.insert(0, "0")
ramp_down_decrement_entry.pack()

# ====== Output Labels ======
tk.Label(window, text="Remaining Time:").pack()
timer_label = tk.Label(window, text="0 s", fg="blue", font=("Arial", 14))
timer_label.pack()

tk.Label(window, text="Current Value (%):").pack()
current_value_label = tk.Label(window, text="0%", fg="green", font=("Arial", 14))
current_value_label.pack()

tk.Label(window, text="Final Hold Countdown:").pack()
final_hold_timer_label = tk.Label(window, text="0:00", fg="purple", font=("Arial", 14))
final_hold_timer_label.pack()

# ====== Globals ======
stop_signal = threading.Event()
force_ramp_down = threading.Event()
serial_conn = None
log_data = []

# ====== Emergency Stop ======
def emergency_stop():
    stop_signal.set()
    force_ramp_down.set()
    try:
        if serial_conn and serial_conn.is_open:
            serial_conn.write(b"0\n")
            serial_conn.close()
    except Exception as e:
        print("Emergency stop failed to send 0:", e)
    current_value_label.config(text="0%")
    timer_label.config(text="0 s")
    final_hold_timer_label.config(text="0:00")
    messagebox.showerror("EMERGENCY STOP", "Emergency stop triggered! System halted.")

# ====== Validation Check ======
def validate_inputs():
    if not port_combo.get():
        raise ValueError("Please select a port.")
    try:
        start = int(start_entry.get())
        end = int(end_entry.get())
        increment = int(increment_entry.get())
        delay = int(delay_entry.get())
    except ValueError:
        raise ValueError("All values must be whole numbers (integers).")

    if not (0 <= start <= 100):
        raise ValueError("Starting point must be between 0 and 100.")
    if not (0 <= end <= 100):
        raise ValueError("Ending point must be greater than starting point.")
    if end <= start:
        raise ValueError("Ending point must be greater than starting point.")
    if increment <= 0:
        raise ValueError("Increment must be a positive integer.")
    if delay <= 0:
        raise ValueError("Time delay must be greater than 0.")
    return port_combo.get(), start, end, increment, delay

# ====== Core Logic ======
def connect_serial(port):
    try:
        ser = serial.Serial(port, 9600, timeout=1)
        time.sleep(2)
        return ser
    except serial.SerialException:
        raise ValueError(f"Cannot open port {port}. Make sure the device is connected.")

def parse_telemetry(response):
    try:
        data = {'rms': None, 'weight': None, 'thrust': None, 'prm': None}
        parts = response.split(',')
        for part in parts:
            key_val = part.strip().split(':')
            if len(key_val) == 2:
                key = key_val[0].lower().strip()
                val = key_val[1].strip().lower()
                if key == 'rms':
                    data['rms'] = float(val)
                elif key == 'weight':
                    data['weight'] = float(val.replace('g', ''))
                elif key == 'thrust':
                    data['thrust'] = float(val.replace('n', ''))
                elif key == 'prm':
                    data['prm'] = int(val)
        return data
    except Exception:
        return {'rms': None, 'weight': None, 'thrust': None, 'prm': None}

def save_to_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "TestBench Data"
    ws.append(["Timestamp", "Sent (%)", "RMS", "Weight (g)", "Thrust (N)", "PRM"])
    for row in log_data:
        ws.append(row)
    filename = f"test_bench_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(filename)
    messagebox.showinfo("Saved", f"Data saved to {filename}")

def generate_steps(start, end, increment):
    steps = []
    current = start
    while current <= end:
        steps.append(current)
        current += increment
    if steps[-1] > end:
        steps[-1] = end
    elif steps[-1] != end:
        steps.append(end)
    return steps

def send_and_log_step(current, delay_seconds):
    if stop_signal.is_set():
        return

    if serial_conn:
        serial_conn.write(f"{current}\n".encode())

    response = ""
    if serial_conn.in_waiting:
        try:
            response = serial_conn.readline().decode(errors='ignore').strip()
        except:
            response = ""

    telemetry = parse_telemetry(response)
    log_data.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        current,
        telemetry.get('rms'),
        telemetry.get('weight'),
        telemetry.get('thrust'),
        telemetry.get('prm')
    ])

    current_value_label.config(text=f"{current}%")
    window.update()

    for remaining in range(delay_seconds, 0, -1):
        if stop_signal.is_set():
            break
        timer_label.config(text=f"{remaining} s")
        window.update()
        time.sleep(1)

    timer_label.config(text="0 s")

def run_simulation():
    global serial_conn, log_data
    log_data = []

    try:
        port, start, end, increment, delay = validate_inputs()
        serial_conn = connect_serial(port)
        try:
            final_hold_minutes = int(final_hold_entry.get())
            ramp_down_delay = int(ramp_down_delay_entry.get())
            ramp_down_decrement = int(ramp_down_decrement_entry.get())
            if final_hold_minutes < 0 or ramp_down_delay <= 0 or ramp_down_decrement <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Input Error", "Final Hold, Ramp-Down Delay, and Ramp-Down Decrement must be positive integers.")
            return
    except ValueError as ve:
        messagebox.showerror("Input Error", str(ve))
        return

    steps = generate_steps(start, end, increment)
    remaining_time = len(steps) * delay

    def update_timer():
        nonlocal remaining_time
        while remaining_time > 0 and not stop_signal.is_set():
            time.sleep(1)
            remaining_time -= 1

    threading.Thread(target=update_timer, daemon=True).start()

    current = steps[0]
    for step in steps:
        if force_ramp_down.is_set() or stop_signal.is_set():
            break
        current = step
        if step == steps[-1] and final_hold_minutes > 0 and not stop_signal.is_set():
            send_and_log_step(step, 0)
            total_seconds = final_hold_minutes * 60
            while total_seconds > 0 and not stop_signal.is_set() and not force_ramp_down.is_set():
                mins, secs = divmod(total_seconds, 60)
                final_hold_timer_label.config(text=f"{mins}:{secs:02}")
                window.update()
                time.sleep(1)
                total_seconds -= 1
            final_hold_timer_label.config(text="0:00")
        else:
            send_and_log_step(step, delay)

    if stop_signal.is_set():
        return

    ramp_start_value = current
    reverse_steps = []
    temp = ramp_start_value
    while temp >= 0:
        reverse_steps.append(temp)
        temp -= ramp_down_decrement
    if reverse_steps[-1] != 0:
        reverse_steps.append(0)

    for val in reverse_steps:
        if stop_signal.is_set():
            break
        send_and_log_step(val, ramp_down_delay)

    current_value_label.config(text="0%")
    timer_label.config(text="0 s")

    if serial_conn:
        serial_conn.close()

    if not stop_signal.is_set():
        save_to_excel()

    stop_signal.set()

# ====== Buttons ======
def start_process():
    stop_signal.clear()
    force_ramp_down.clear()
    threading.Thread(target=run_simulation, daemon=True).start()

def stop_process():
    if not force_ramp_down.is_set():
        force_ramp_down.set()
        messagebox.showinfo("Stopping", "Ramp-down will begin from current value.")

button_frame = tk.Frame(window)
button_frame.pack(pady=20)

start_btn = tk.Button(button_frame, text="Start", bg="lightgreen",
                      command=start_process, width=30, height=2)
start_btn.pack(pady=5)

end_btn = tk.Button(button_frame, text="End", bg="tomato",
                    command=stop_process, width=30, height=2)
end_btn.pack(pady=5)

emergency_btn = tk.Button(button_frame, text="EMERGENCY STOP", bg="red", fg="white",
                          font=("Arial", 16, "bold"), width=20, height=3,
                          command=emergency_stop)
emergency_btn.pack(pady=10)



refresh_ports()
window.mainloop()