import random, sqlite3, datetime, os, threading
import tkinter as tk
from tkinter import messagebox, ttk
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import requests
import psutil

# ==============================
# USERS
# ==============================
USERS = {
    "admin": {"password": "admin123", "role": "admin"},
    "user": {"password": "user123", "role": "user"}
}

# ==============================
# DATABASE
# ==============================
conn = sqlite3.connect("cloud.db")
cur = conn.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS resources (
    name TEXT, cpu INTEGER, memory INTEGER,
    storage INTEGER, network INTEGER,
    cost REAL, risk TEXT, sla TEXT,
    action TEXT, timestamp TEXT
)
""")
conn.commit()

# ==============================
# MAIN APP
# ==============================
class CloudApp:
    def __init__(self, root, role):
        self.root = root
        self.role = role
        self.root.title("☁ Cloud Monitoring Dashboard")
        self.root.state('zoomed')   # Maximized window
        self.root.configure(bg="#1e1e2f")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)

        self.dashboard_tab = tk.Frame(self.notebook, bg="#1e1e2f")
        self.graph_tab = tk.Frame(self.notebook, bg="#1e1e2f")
        self.report_tab = tk.Frame(self.notebook, bg="#1e1e2f")

        self.notebook.add(self.dashboard_tab, text="Dashboard")
        self.notebook.add(self.graph_tab, text="Live Graph")
        self.notebook.add(self.report_tab, text="Reports")

        self.setup_dashboard()
        self.setup_graph()
        self.setup_reports()

        self.start_live_monitoring()

    # ================= DASHBOARD =================
    def setup_dashboard(self):
        frame = self.dashboard_tab

        self.resource_entry = tk.Entry(frame, font=("Arial", 12))
        self.resource_entry.pack(pady=10)

        # Add Resource (both can use)
        tk.Button(frame, text="➕ Add Resource", command=self.manual_add).pack(pady=5)

        # Summary (both can use)
        tk.Button(frame, text="📊 Show Summary", command=self.show_summary).pack(pady=5)
        tk.Button(frame, text="🚪 Logout", command=self.logout, bg="red", fg="white").pack(pady=10)
        # ❗ ONLY ADMIN
        if self.role == "admin":
            tk.Button(frame, text="🗑 Reset DB", command=self.reset_database).pack(pady=5)

        # TABLE
        self.table = ttk.Treeview(frame, columns=("Name","CPU","Memory","Cost"), show="headings")
        for col in ("Name","CPU","Memory","Cost"):
            self.table.heading(col, text=col)
        self.table.pack(fill="both", expand=True, pady=10)

        self.log_box = tk.Text(frame, height=6, bg="black", fg="lime")
        self.log_box.pack(fill="x")

    # ================= GRAPH =================
    def setup_graph(self):
        self.fig = Figure(figsize=(6,4))
        self.ax = self.fig.add_subplot(111)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_tab)
        self.canvas.get_tk_widget().pack()

        self.cpu_data = []
        self.mem_data = []

        self.update_graph()

    def update_graph(self):
        cpu = psutil.cpu_percent()
        mem = psutil.virtual_memory().percent

        self.cpu_data.append(cpu)
        self.mem_data.append(mem)

        self.ax.clear()
        self.ax.plot(self.cpu_data, label="CPU %")
        self.ax.plot(self.mem_data, label="Memory %")
        self.ax.legend()
        self.ax.set_title("Live Monitoring")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Usage Percentage")

        # 🚨 ALERT CONDITIONS
        if cpu > 80:
            self.show_alert("⚠ High CPU Usage!")

        if mem > 80:
            self.show_alert("⚠ High Memory Usage!")

        self.canvas.draw()
        self.load_table()
        self.root.after(1000, self.update_graph)

    # ================= REPORT =================
    def setup_reports(self):
        frame = self.report_tab

        tk.Button(frame, text="📁 Export Excel", command=self.export_excel).pack(pady=5)

        # Charts for both
        tk.Button(frame, text="🥧 Pie Chart", command=self.show_pie_chart).pack(pady=5)
        tk.Button(frame, text="📊 Bar Chart", command=self.show_bar_chart).pack(pady=5)
        tk.Button(frame, text="📈 Cost Trend", command=self.show_trend_chart).pack(pady=5)
        tk.Button(frame, text="🤖 AI Prediction", command=self.show_prediction_chart).pack(pady=5)

        self.chart_frame = tk.Frame(frame, bg="#1e1e2f")
        self.chart_frame.pack(fill="both", expand=True)
    # ================= ADD RESOURCE =================
    def manual_add(self):
        name = self.resource_entry.get()
        if name:
            self.add_resource(name)

    def add_resource(self, name):
        cpu = random.randint(10,90)
        memory = random.randint(10,90)
        storage = random.randint(10,90)
        network = random.randint(10,90)
        cost = round(random.uniform(10,100),2)
        self.detect_anomaly(cost)

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cur.execute("INSERT INTO resources VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (name,cpu,memory,storage,network,cost,"Low","OK","Monitor",timestamp))
        conn.commit()

        self.log_box.insert("end", f"\nAdded: {name}")
        self.load_table()

        self.send_to_ollama(f"Analyze CPU {cpu} and Memory {memory}")

    # ================= TABLE =================
    def load_table(self):
        for i in self.table.get_children():
            self.table.delete(i)

        for row in cur.execute("SELECT name,cpu,memory,cost FROM resources"):
            self.table.insert("", "end", values=row)

    # ================= EXCEL EXPORT =================
    def export_excel(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Cloud Report"

        headers = ["Name","CPU","Memory","Storage","Network","Cost","Time"]
        ws.append(headers)

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        for row in cur.execute("SELECT name,cpu,memory,storage,network,cost,timestamp FROM resources"):
            ws.append(row)

        wb.save("Cloud_Report.xlsx")

        messagebox.showinfo("Success", "Excel Report Generated!")

    # ================= SUMMARY =================
    def show_summary(self):
        total, count = 0, 0
        for row in cur.execute("SELECT cost FROM resources"):
            total += row[0]
            count += 1

        avg = total/count if count else 0

        messagebox.showinfo("Summary",
                            f"Resources: {count}\nAvg Cost: {round(avg,2)}")
    def logout(self):
        self.root.destroy()
        show_login()

    def show_alert(self, message):
        self.log_box.insert("end", f"\n{message}")
        self.log_box.see("end")

    def detect_anomaly(self, new_cost):
        costs = []

        for row in cur.execute("SELECT cost FROM resources"):
            costs.append(row[0])

        if len(costs) < 5:
            return

        avg = sum(costs) / len(costs)

        if new_cost > 2 * avg:
            self.show_alert(f"🚨 Cost Anomaly Detected! Cost: {new_cost}")   
    # ================= RESET =================
    def reset_database(self):
        cur.execute("DELETE FROM resources")
        conn.commit()
        self.load_table()
        self.log_box.insert("end","\nDatabase Reset")

    # ================= OLLAMA =================
    def _call_ollama(self, prompt):
        try:
            r = requests.post(
                "http://localhost:11434/api/generate",
                json={
                    "model": "tinyllama",
                    "prompt": prompt,
                    "stream": False
                }
            )
            res = r.json().get("response", "")

            # ✅ SAFE UI UPDATE
            self.root.after(0, lambda: self.log_box.insert("end", f"\n[AI]: {res}"))

        except Exception as e:
            self.root.after(0, lambda: self.log_box.insert("end", f"\nAI Error: {e}"))

    def send_to_ollama(self, prompt):
        threading.Thread(
            target=self._call_ollama,
            args=(prompt,),
            daemon=True
        ).start()
    
    
    # ================= AUTO =================
    def start_live_monitoring(self):
        for _ in range(2):
            self.add_resource("Auto")

    def clear_chart(self):
        for widget in self.chart_frame.winfo_children():
            widget.destroy()

    def draw_chart(self, fig):
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()

    def show_bar_chart(self):
        self.clear_chart()

        names, cpu, memory, cost = [], [], [], []

        for row in cur.execute("SELECT name, cpu, memory, cost FROM resources"):
            names.append(row[0])
            cpu.append(row[1])
            memory.append(row[2])
            cost.append(row[3])

        fig = Figure(figsize=(6,4))
        ax = fig.add_subplot(111)

        x = range(len(names))
        ax.bar(x, cost)
        ax.set_xticks(x)
        ax.set_xticklabels(names, rotation=45)
        ax.set_title("Cost per Resource")
        ax.set_xlabel("Resources Name")
        ax.set_ylabel("Cost")

        self.draw_chart(fig)

    def show_trend_chart(self):
        self.clear_chart()

        costs = []

        for row in cur.execute("SELECT cost FROM resources"):
            costs.append(row[0])

        fig = Figure(figsize=(6,4))
        ax = fig.add_subplot(111)

        ax.plot(costs, marker='o')
        ax.set_title("Cost Trend Over Time")
        ax.set_xlabel("Entries")
        ax.set_ylabel("Cost")
        self.draw_chart(fig)

    def show_pie_chart(self):
        self.clear_chart()   # clear previous chart

        labels = []
        costs = []

        for row in cur.execute("SELECT name, cost FROM resources"):
            labels.append(row[0])
            costs.append(row[1])

        if not costs:
            messagebox.showinfo("Info", "No data available")
            return

        fig = Figure(figsize=(6,4))
        ax = fig.add_subplot(111)

        ax.pie(costs, labels=labels, autopct='%1.1f%%')
        ax.set_title("Cost Distribution")

        self.draw_chart(fig)

    def show_prediction_chart(self):
        self.clear_chart()

        costs = []

        for row in cur.execute("SELECT cost FROM resources"):
            costs.append(row[0])

        if len(costs) < 2:
            messagebox.showinfo("Info", "Not enough data")
            return

        x = list(range(len(costs)))

        n = len(x)
        avg_x = sum(x)/n
        avg_y = sum(costs)/n

        num = sum((x[i]-avg_x)*(costs[i]-avg_y) for i in range(n))
        den = sum((x[i]-avg_x)**2 for i in range(n))
        slope = num/den if den != 0 else 0

        intercept = avg_y - slope*avg_x

        future_x = list(range(len(costs)+5))
        predicted = [slope*i + intercept for i in future_x]

        fig = Figure(figsize=(6,4))
        ax = fig.add_subplot(111)

        ax.plot(x, costs, label="Actual", marker='o')
        ax.plot(future_x, predicted, linestyle='dashed', label="Predicted")

        ax.legend()
        ax.set_title("AI Cost Prediction")
        ax.set_xlabel("Entities")
        ax.set_ylabel("Cost")

        self.draw_chart(fig)

# ================= LOGIN =================
def login():
    u = user_entry.get()
    p = pass_entry.get()

    if u in USERS and USERS[u]["password"] == p:
        login_win.destroy()
        root = tk.Tk()
        CloudApp(root, USERS[u]["role"])
        root.mainloop()
    else:
        messagebox.showerror("Error","Invalid Login")

def show_login():
    global login_win, user_entry, pass_entry

    login_win = tk.Tk()
    login_win.title("Login")
    login_win.state('zoomed')

    tk.Label(login_win,text="Username").pack()
    user_entry = tk.Entry(login_win)
    user_entry.pack()

    tk.Label(login_win,text="Password").pack()
    pass_entry = tk.Entry(login_win, show="*")
    pass_entry.pack()

    tk.Button(login_win,text="Login",command=login).pack()

    login_win.mainloop()

# ================= RUN =================
if __name__ == "__main__":
    show_login()