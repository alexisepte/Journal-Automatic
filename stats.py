import tkinter as tk
from tkinter import ttk, messagebox
import random

def create_mock_data(num_trades=50):
    setups = ["Breakout", "Retest", "Trend Continuation", "Reversal", "Scalp"]
    entry_types = ["Limit", "Market"]
    market_sessions = ["New York", "London", "Asia"]
    sl_reasons = ["Hit SL", "Manual Close (SL)", "Time Expiry"]
    close_reasons = ["Take Profit", "Manual Close (Profit)", "Hit SL", "Manual Close (Loss)"]
    tp_reasons = ["Target Hit", "Manual Close (TP)"]

    data = []
    for i in range(1, num_trades + 1):
        win = random.choice([True, False])
        pnl = round(random.uniform(50, 500) if win else random.uniform(-500, -50), 2)
        r_multiple = round(random.uniform(1, 5) if win else random.uniform(-0.5, -2), 2)
        hold_time = random.randint(15, 240)

        trade = {
            "ID": i,
            "Setup": random.choice(setups),
            "Entry Type": random.choice(entry_types),
            "Market Session": random.choice(market_sessions),
            "Stop Loss Reason": random.choice(sl_reasons) if not win else "N/A",
            "Reason for Close": random.choice(close_reasons),
            "Take Profit Reason": random.choice(tp_reasons) if win else "N/A",
            "Stop Loss Size": random.uniform(0.5, 2.0),
            "P&L": pnl,
            "R:R": r_multiple,
            "Win": win,
            "Hold Time": hold_time
        }
        data.append(trade)
    return data

class StatsPage(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Playbook Stats Page")
        self.geometry("1200x800")
        self.protocol("WM_DELETE_WINDOW", self.quit)
        self.data = create_mock_data()
        self.filtered_data = self.data[:]
        self.create_widgets()
        self.apply_filters()

    def create_widgets(self):
        filter_frame = ttk.LabelFrame(self, text="Smart Filter Panel", padding="10")
        filter_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        filter_options = {
            "Setup": sorted(list(set(d["Setup"] for d in self.data))),
            "Entry Type": sorted(list(set(d["Entry Type"] for d in self.data))),
            "Market Session": sorted(list(set(d["Market Session"] for d in self.data))),
            "Stop Loss Reason": sorted(list(set(d["Stop Loss Reason"] for d in self.data))),
            "Reason for Close": sorted(list(set(d["Reason for Close"] for d in self.data))),
            "Take Profit Reason": sorted(list(set(d["Take Profit Reason"] for d in self.data))),
        }

        self.filter_vars = {}
        row = col = 0
        for label_text, options in filter_options.items():
            ttk.Label(filter_frame, text=f"{label_text}:").grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
            var = ttk.Combobox(filter_frame, values=["ANY"] + options, state="readonly")
            var.set("ANY")
            var.grid(row=row, column=col+1, sticky=tk.EW, padx=5, pady=2)
            self.filter_vars[label_text] = var
            col += 2
            if col >= 6:
                row += 1
                col = 0

        ttk.Label(filter_frame, text="Min SL Size (%):").grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
        self.filter_vars["Stop Loss Size"] = ttk.Entry(filter_frame, width=10)
        self.filter_vars["Stop Loss Size"].grid(row=row, column=col+1, sticky=tk.EW, padx=5, pady=2)
        ttk.Button(filter_frame, text="Apply Filters", command=self.apply_filters).grid(row=row, column=col+2, padx=10, pady=5)

        stats_frame = ttk.LabelFrame(self, text="Stats Overview (After Filtering)", padding="10")
        stats_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        self.stats_labels = {}
        stats_data = {
            "Total Trades": "0", "Win Rate %": "0.0%", "Avg. R:R": "0.00",
            "Total P&L": "$0.00", "# of Wins and Losses": "0 Wins / 0 Losses",
            "Avg. Hold Time": "0 min"
        }
        for i, (label_text, default_value) in enumerate(stats_data.items()):
            ttk.Label(stats_frame, text=f"{label_text}:").grid(row=0 if i < 3 else 1, column=i%3*2, sticky=tk.W, padx=5, pady=2)
            label = ttk.Label(stats_frame, text=default_value, font=('Arial', 10, 'bold'))
            label.grid(row=0 if i < 3 else 1, column=i%3*2+1, sticky=tk.W, padx=5, pady=2)
            self.stats_labels[label_text] = label

        ttk.Label(stats_frame, text="Win Rate Visual:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.win_rate_canvas = tk.Canvas(stats_frame, width=200, height=20, bg="lightgray", highlightbackground="gray")
        self.win_rate_canvas.grid(row=2, column=1, columnspan=5, sticky=tk.W, padx=5, pady=5)
        self.win_rate_bar = self.win_rate_canvas.create_rectangle(0, 0, 0, 20, fill="red", outline="")
        self.win_rate_text = self.win_rate_canvas.create_text(100, 10, text="", fill="black", font=('Arial', 9, 'bold'))

        table_frame = ttk.LabelFrame(self, text="Trades Table", padding="10")
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        columns = ("ID", "Setup", "Entry Type", "P&L", "R:R", "Win", "Hold Time")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c))
            self.tree.column(col, width=100, anchor=tk.CENTER)
        self.tree.column("ID", width=40)
        self.tree.column("Setup", width=120)
        self.tree.column("Entry Type", width=80)
        self.tree.column("P&L", width=90)
        self.tree.column("R:R", width=60)
        self.tree.column("Win", width=50)
        self.tree.column("Hold Time", width=80)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", self.on_trade_select)
        self.tree.bind("<Return>", self.on_trade_select)

    def apply_filters(self):
        current_filters = {key: var.get() for key, var in self.filter_vars.items()}
        self.filtered_data = []
        for trade in self.data:
            match = True
            for filter_name, filter_value in current_filters.items():
                if filter_name == "Stop Loss Size":
                    try:
                        min_sl_size = float(filter_value)
                        if trade["Stop Loss Size"] < min_sl_size:
                            match = False
                            break
                    except ValueError:
                        pass
                else:
                    if filter_value != "ANY" and trade[filter_name] != filter_value:
                        match = False
                        break
            if match:
                self.filtered_data.append(trade)
        self.update_stats_and_table()

    def update_stats_and_table(self):
        total_trades = len(self.filtered_data)
        wins = sum(1 for trade in self.filtered_data if trade["Win"])
        losses = total_trades - wins
        win_rate = (wins / total_trades * 100) if total_trades > 0 else 0.0
        total_pnl = sum(trade["P&L"] for trade in self.filtered_data)
        avg_rr = sum(trade["R:R"] for trade in self.filtered_data) / total_trades if total_trades > 0 else 0.0
        avg_hold_time = sum(trade["Hold Time"] for trade in self.filtered_data) / total_trades if total_trades > 0 else 0.0

        self.stats_labels["Total Trades"].config(text=str(total_trades))
        self.stats_labels["Win Rate %"].config(text=f"{win_rate:.1f}%")
        self.stats_labels["Avg. R:R"].config(text=f"{avg_rr:.2f}")
        self.stats_labels["Total P&L"].config(text=f"${total_pnl:.2f}")
        self.stats_labels["# of Wins and Losses"].config(text=f"{wins} Wins / {losses} Losses")
        self.stats_labels["Avg. Hold Time"].config(text=f"{avg_hold_time:.0f} min")

        bar_width = (win_rate / 100) * 200
        color = "red"
        if win_rate > 80:
            color = "green"
        elif win_rate > 50:
            color = "yellow"
        self.win_rate_canvas.coords(self.win_rate_bar, 0, 0, bar_width, 20)
        self.win_rate_canvas.itemconfig(self.win_rate_bar, fill=color)
        self.win_rate_canvas.itemconfig(self.win_rate_text, text=f"{win_rate:.1f}%")
        self.win_rate_canvas.coords(self.win_rate_text, bar_width / 2, 10)

        for item in self.tree.get_children():
            self.tree.delete(item)
        for trade in self.filtered_data:
            self.tree.insert("", "end", values=(
                trade["ID"], trade["Setup"], trade["Entry Type"],
                f"${trade['P&L']:.2f}", f"{trade['R:R']:.2f}",
                "✅" if trade["Win"] else "❌", f"{trade['Hold Time']} min"
            ))

    def on_trade_select(self, event):
        selected_item = self.tree.focus()
        if selected_item:
            trade_id = self.tree.item(selected_item, "values")[0]
            messagebox.showinfo("Trade Details", f"Opening journal entry for Trade ID: {trade_id}\n\n(This would be a popup with full trade details)")

    def sort_column(self, col):
        messagebox.showinfo("Sorting", f"Sorting by column: {col}\n\n(Implement actual sorting logic here)")

if __name__ == "__main__":
    app = StatsPage()
    app.mainloop()
