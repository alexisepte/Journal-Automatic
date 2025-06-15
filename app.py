import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime, time as dt_time, timedelta
import pytz
import os
import json

try:
    from PIL import Image, ImageTk
except ImportError:
    import sys
    messagebox.showerror("Missing Dependency", "Pillow (PIL) is required for image thumbnails.\nInstall it with:\n\npip install pillow")
    sys.exit(1)

# --- Constants (These will now be initial defaults and then loaded from files) ---
PIP_VALUE_XAUUSD = 0.1
USD_PER_PIP_PER_LOT = 10

# Default lists - these will be used to create the initial JSON files if they don't exist.
DEFAULT_SL_LOGIC = ["Below Support", "ATR Stop", "Structure", "Other"]
DEFAULT_TP_LOGIC = ["At Resistance", "RR Ratio", "Previous High", "Other"]
DEFAULT_SETUPS = ["Breakout", "Reversal", "Pullback", "Trend Continuation", "Range", "News Play", "Other"]
DEFAULT_ENTRIES = ["Market", "Limit", "Stop", "Break-Even", "Retest", "Other"]
DEFAULT_PARTIAL_CLOSE_REASONS = [
    "", # Default empty option
    "Reached Partial TP 1",
    "Reached Partial TP 2",
    "Minor Support/Resistance Hit",
    "Candle Closed Against Me",
    "Volatility Spike",
    "News Event Approaching",
    "Time Based Exit",
    "Price Action Shift",
    "Manual Intervention",
    "Other"
]

TIMEFRAME_ENTRIES = ["15m", "30m", "1h", "4h", "1d"]
MARKET_SESSIONS_UTC = [
    ("Sydney", dt_time(21, 0), dt_time(6, 0)),
    ("Tokyo", dt_time(0, 0), dt_time(9, 0)),
    ("London", dt_time(8, 0), dt_time(17, 0)),
    ("New York", dt_time(13, 0), dt_time(22, 0)),
]
COMMON_TIMEZONES = ["UTC", "US/Eastern", "Europe/London", "Asia/Tokyo", "Australia/Sydney"]

TRADES_FILE = "trades_journal.json" # File where all journal data will be stored

# File paths for playbook options
PLAYBOOK_DIR = "playbook_data"
SETUP_FILE = os.path.join(PLAYBOOK_DIR, "setups.json")
ENTRY_FILE = os.path.join(PLAYBOOK_DIR, "entries.json")
SL_REASONS_FILE = os.path.join(PLAYBOOK_DIR, "sl_reasons.json")
TP_REASONS_FILE = os.path.join(PLAYBOOK_DIR, "tp_reasons.json")
PARTIAL_CLOSE_REASONS_FILE = os.path.join(PLAYBOOK_DIR, "close_reasons.json")

# --- Trade Class (No change, as it uses dictionaries for info) ---
class Trade:
    def __init__(self, symbol, timeframe, info=None, tf_screenshots=None, review=None, partial_closes=None, sl_to_be=False):
        self.symbol = symbol
        self.timeframe = timeframe
        self.info = info or {}
        self.tf_screenshots = tf_screenshots or {
            "D1": {"before": None, "after": None},
            "H4": {"before": None, "after": None},
            "H1": {"before": None, "after": None}
        }
        self.review = review or {
            "outcome": "",
            "price": "",
            "notes": "",
            "exit_time": "",
            "max_drawdown_pips": "",
        }
        self.partial_closes = partial_closes or []
        self.sl_to_be = sl_to_be

    def to_dict(self):
        return {
            "symbol": self.symbol,
            "timeframe": self.timeframe,
            "info": self.info,
            "tf_screenshots": self.tf_screenshots,
            "review": self.review,
            "partial_closes": self.partial_closes,
            "sl_to_be": self.sl_to_be
        }

    @classmethod
    def from_dict(cls, d):
        default_review = {
            "outcome": "",
            "price": "",
            "notes": "",
            "exit_time": "",
            "max_drawdown_pips": "",
        }
        review_data = d.get("review", {})
        review_data = {**default_review, **review_data}

        default_tf_screenshots = {
            "D1": {"before": None, "after": None},
            "H4": {"before": None, "after": None},
            "H1": {"before": None, "after": None}
        }
        tf_screenshots_data = d.get("tf_screenshots", {})
        for tf in default_tf_screenshots:
            if tf not in tf_screenshots_data:
                tf_screenshots_data[tf] = default_tf_screenshots[tf]
            else:
                if "before" not in tf_screenshots_data[tf]:
                    tf_screenshots_data[tf]["before"] = None
                if "after" not in tf_screenshots_data[tf]:
                    tf_screenshots_data[tf]["after"] = None

        partial_closes_data = d.get("partial_closes", [])
        for pc in partial_closes_data:
            pc.setdefault("pips", 0.0)
            if "reason_for_close" not in pc:
                if "notes" in pc and pc["notes"] is not None:
                    pc["reason_for_close"] = pc["notes"]
                else:
                    pc["reason_for_close"] = ""
                pc.pop("notes", None)
            pc.setdefault("pnl", 0.0)
        
        sl_to_be_val = d.get("sl_to_be", False)

        return cls(
            d.get("symbol", ""),
            d.get("timeframe", ""),
            d.get("info", {}),
            tf_screenshots_data,
            review_data,
            partial_closes_data,
            sl_to_be=sl_to_be_val
        )

# --- TradingJournalApp Class ---
class TradingJournalApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Trading Journal")
        self.geometry("1500x780")
        self.trades = []
        self.account_balance_var = tk.DoubleVar(value=10000.00)
        self.trade_stats_var = tk.StringVar()
        self._sl_tp_is_updating = False

        # Load playbook options
        self.load_playbook_options()

        self.tab_control = ttk.Notebook(self)
        self.tab_control.pack(fill="both", expand=True)

        # Statistics and Balance Display
        stats_frame = ttk.Frame(self)
        stats_frame.pack(fill="x", padx=0, pady=0)
        self.stats_label = ttk.Label(stats_frame, textvariable=self.trade_stats_var, font=("Segoe UI", 10, "bold"), foreground="#444")
        self.stats_label.pack(side="left", padx=(14, 5), pady=(5, 0))
        bal_label = ttk.Label(stats_frame, text="Account Balance:", font=("Segoe UI", 10, "bold"))
        bal_label.pack(side="left", padx=(10,2), pady=(5,0))
        self.bal_disp = ttk.Label(stats_frame, textvariable=self.account_balance_var, font=("Segoe UI", 11, "bold"), foreground="#228B22")
        self.bal_disp.pack(side="left", padx=(3, 0), pady=(5,0))
        update_bal_btn = ttk.Button(stats_frame, text="Update Balance", command=self.update_balance_popup, width=14)
        update_bal_btn.pack(side="right", padx=(0, 8), pady=(5,0))
        self._on_focus_select_all(update_bal_btn, self.account_balance_var)

        # Journal Entry Tab
        self.journal_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.journal_tab, text="Journal Entry")
        self.build_journal_tab()

        # Journal Review Tab
        self.review_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.review_tab, text="Journal Review")
        self.build_review_tab()

        # Playbook Tab (NEW)
        self.playbook_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.playbook_tab, text="Playbook")
        self.build_playbook_tab() # Call the new method

        # Menubar
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Save", command=self.save_trades)
        filemenu.add_command(label="Load", command=self.load_trades)
        menubar.add_cascade(label="File", menu=filemenu)
        self.config(menu=menubar)

        # Initial load and update
        self.load_trades()
        self.update_stats_bar()

    def _on_focus_select_all(self, widget, tk_var_ref=None):
        """Binds a function to a widget's <FocusIn> event to select all text."""
        def select_all(event):
            event.widget.select_range(0, tk.END)
            event.widget.icursor(tk.END)
        widget.bind("<FocusIn>", select_all)

    def load_playbook_options(self):
        """Loads playbook options from JSON files or creates them with defaults."""
        os.makedirs(PLAYBOOK_DIR, exist_ok=True) # Ensure directory exists

        # Helper function to load/create a file
        def load_or_create(filepath, default_list):
            if os.path.exists(filepath):
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        return json.load(f)
                except json.JSONDecodeError:
                    messagebox.showwarning("File Error", f"Corrupted playbook file: {filepath}. Resetting to default.")
                    # Overwrite corrupted file with defaults
                    with open(filepath, 'w', encoding='utf-8') as f:
                        json.dump(default_list, f, indent=2)
                    return default_list
            else:
                # Create file with defaults if it doesn't exist
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(default_list, f, indent=2)
                return default_list

        self.sl_logic_options = load_or_create(SL_REASONS_FILE, DEFAULT_SL_LOGIC)
        self.tp_logic_options = load_or_create(TP_REASONS_FILE, DEFAULT_TP_LOGIC)
        self.setups_options = load_or_create(SETUP_FILE, DEFAULT_SETUPS)
        self.entries_options = load_or_create(ENTRY_FILE, DEFAULT_ENTRIES)
        self.partial_close_reasons_options = load_or_create(PARTIAL_CLOSE_REASONS_FILE, DEFAULT_PARTIAL_CLOSE_REASONS)

    def save_playbook_options(self, options_list, filepath):
        """Saves a list of playbook options to a JSON file."""
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(options_list, f, indent=2)
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save options to {filepath}: {e}")

    def refresh_journal_dropdowns(self):
        """Refreshes the combobox dropdowns on the Journal Entry tab."""
        # This method will be called after changes are made in the Playbook tab
        self.sl_reason_combo["values"] = self.sl_logic_options
        self.tp_reason_combo["values"] = self.tp_logic_options
        self.setup_combo["values"] = self.setups_options
        self.entry_combo["values"] = self.entries_options
        # Partial close reason combo is in the review popup, so it needs to be updated when that opens
        # We can't directly update it here unless it's a global variable or part of a persistent frame.
        # For now, it will refresh when the review popup is opened.

    def build_journal_tab(self):
        # Tkinter Variable definitions
        self.trade_type_var = tk.StringVar(value="Buy")
        self.symbol_var = tk.StringVar(value="XAUUSD")
        self.trade_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.trade_time_var = tk.StringVar(value=datetime.now().strftime("%H:%M"))
        self.market_session_var = tk.StringVar(value="")
        self.timezone_var = tk.StringVar(value="UTC")
        self.entry_price_var = tk.DoubleVar(value=2300.0)
        self.lot_size_var = tk.DoubleVar(value=1.0)
        self.sl_pips_var = tk.DoubleVar()
        self.sl_price_var = tk.DoubleVar()
        self.sl_loss_var = tk.StringVar(value="")
        self.sl_loss_pct_var = tk.StringVar(value="")
        self.sl_reason_var = tk.StringVar()
        self.tp_pips_var = tk.DoubleVar()
        self.tp_price_var = tk.DoubleVar()
        self.tp_profit_var = tk.StringVar(value="")
        self.tp_profit_pct_var = tk.StringVar(value="")
        self.tp_reason_var = tk.StringVar()
        self.setup_var = tk.StringVar()
        self.entry_var = tk.StringVar()
        self.timeframe_var = tk.StringVar()
        self.timezone_display_var = tk.StringVar(value="UTC")

        # Top Frame: Trade Type & Symbol
        top_frame = ttk.Frame(self.journal_tab)
        top_frame.pack(padx=10, pady=5, fill="x")

        btn_frame = tk.Frame(top_frame)
        btn_frame.pack(side="left", padx=(0,12))
        self.buy_btn = tk.Button(
            btn_frame, text="Buy", width=8, height=2, relief="flat",
            command=self.select_buy, cursor="hand2"
        )
        self.sell_btn = tk.Button(
            btn_frame, text="Sell", width=8, height=2, relief="flat",
            command=self.select_sell, cursor="hand2"
        )
        self.buy_btn.grid(row=0, column=0, padx=(0,2))
        self.sell_btn.grid(row=0, column=1, padx=(2,0))
        self._update_trade_type_buttons()

        ttk.Label(top_frame, text="Symbol:").pack(side="left", padx=10)
        ttk.Entry(top_frame, textvariable=self.symbol_var, width=10).pack(side="left")

        account_frame = ttk.Frame(top_frame)
        account_frame.pack(side="right", padx=10)
        tz_label = ttk.Label(account_frame, text="Time Zone:")
        tz_label.grid(row=1, column=0, sticky="e", pady=(4,0))
        self.tz_combo = ttk.Combobox(
            account_frame, textvariable=self.timezone_var, values=COMMON_TIMEZONES, width=15, state="readonly"
        )
        self.tz_combo.grid(row=1, column=1, sticky="w", padx=(5, 0), pady=(4,0))

        # Date, Time, Market Session
        row_frame = ttk.Frame(self.journal_tab)
        row_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(row_frame, text="Trade Date:").pack(side="left", padx=(0,2))
        ttk.Entry(row_frame, textvariable=self.trade_date_var, width=10).pack(side="left")
        ttk.Label(row_frame, text="Trade Time:").pack(side="left", padx=(10,2))
        ttk.Entry(row_frame, textvariable=self.trade_time_var, width=6).pack(side="left")
        ttk.Label(row_frame, text="Market Session:").pack(side="left", padx=(10,2))
        self.market_session_lbl = ttk.Label(row_frame, textvariable=self.market_session_var, foreground="blue")
        self.market_session_lbl.pack(side="left")

        # Entry Price & Lot Size
        entry_frame = ttk.Frame(self.journal_tab)
        entry_frame.pack(padx=10, pady=10, fill="x")
        ttk.Label(entry_frame, text="Entry Price:").pack(side="left")
        entry_entry = ttk.Entry(entry_frame, textvariable=self.entry_price_var, width=8)
        entry_entry.pack(side="left", padx=2)
        self._on_focus_select_all(entry_entry, self.entry_price_var)

        ttk.Label(entry_frame, text="Lot Size:").pack(side="left", padx=(10,2))
        lot_entry = ttk.Entry(entry_frame, textvariable=self.lot_size_var, width=6)
        lot_entry.pack(side="left")
        self._on_focus_select_all(lot_entry, self.lot_size_var)

        # Risk Management Frame (Stop Loss & Take Profit)
        risk_frame = ttk.LabelFrame(self.journal_tab, text="Risk Management")
        risk_frame.pack(padx=10, pady=3, fill="x")

        # Stop Loss details
        sl_frame = ttk.Frame(risk_frame)
        sl_frame.pack(fill="x", padx=4, pady=2)
        ttk.Label(sl_frame, text="Stop Loss:").pack(side="left")
        ttk.Label(sl_frame, text="Pips:").pack(side="left", padx=(4,2))
        sl_pips_entry = ttk.Entry(sl_frame, textvariable=self.sl_pips_var, width=8)
        sl_pips_entry.pack(side="left")
        self._on_focus_select_all(sl_pips_entry, self.sl_pips_var)

        ttk.Label(sl_frame, text="Price:").pack(side="left", padx=(8,2))
        sl_price_entry = ttk.Entry(sl_frame, textvariable=self.sl_price_var, width=8)
        sl_price_entry.pack(side="left")
        self._on_focus_select_all(sl_price_entry, self.sl_price_var)

        sl_loss_label = ttk.Label(sl_frame, textvariable=self.sl_loss_var, foreground="red")
        sl_loss_label.pack(side="left", padx=(10,2))
        sl_loss_pct_label = ttk.Label(sl_frame, textvariable=self.sl_loss_pct_var, foreground="red")
        sl_loss_pct_label.pack(side="left", padx=(2,5))
        ttk.Label(sl_frame, text="Reason:").pack(side="left", padx=(10,2))
        # Updated to use self.sl_logic_options
        self.sl_reason_combo = ttk.Combobox(sl_frame, textvariable=self.sl_reason_var, values=self.sl_logic_options, width=14, state="readonly")
        self.sl_reason_combo.pack(side="left")

        # Take Profit details
        tp_frame = ttk.Frame(risk_frame)
        tp_frame.pack(fill="x", padx=4, pady=4)
        ttk.Label(tp_frame, text="Take Profit:").pack(side="left")
        ttk.Label(tp_frame, text="Pips:").pack(side="left", padx=(4,2))
        tp_pips_entry = ttk.Entry(tp_frame, textvariable=self.tp_pips_var, width=8)
        tp_pips_entry.pack(side="left")
        self._on_focus_select_all(tp_pips_entry, self.tp_pips_var)

        ttk.Label(tp_frame, text="Price:").pack(side="left", padx=(8,2))
        tp_price_entry = ttk.Entry(tp_frame, textvariable=self.tp_price_var, width=8)
        tp_price_entry.pack(side="left")
        self._on_focus_select_all(tp_price_entry, self.tp_price_var)

        tp_profit_label = ttk.Label(tp_frame, textvariable=self.tp_profit_var, foreground="green")
        tp_profit_label.pack(side="left", padx=(10,2))
        tp_profit_pct_label = ttk.Label(tp_frame, textvariable=self.tp_profit_pct_var, foreground="green")
        tp_profit_pct_label.pack(side="left", padx=(2,5))
        ttk.Label(tp_frame, text="Reason:").pack(side="left", padx=(10,2))
        # Updated to use self.tp_logic_options
        self.tp_reason_combo = ttk.Combobox(tp_frame, textvariable=self.tp_reason_var, values=self.tp_logic_options, width=14, state="readonly")
        self.tp_reason_combo.pack(side="left")

        # Playbook Frame
        playbook_frame = ttk.LabelFrame(self.journal_tab, text="Playbook")
        playbook_frame.pack(padx=10, pady=10, fill="x")
        ttk.Label(playbook_frame, text="Setup:").pack(side="left", padx=(5,2))
        # Updated to use self.setups_options
        self.setup_combo = ttk.Combobox(playbook_frame, textvariable=self.setup_var, values=self.setups_options, width=16, state="readonly")
        self.setup_combo.pack(side="left", padx=(0,8))
        ttk.Label(playbook_frame, text="Entry:").pack(side="left", padx=(5,2))
        # Updated to use self.entries_options
        self.entry_combo = ttk.Combobox(playbook_frame, textvariable=self.entry_var, values=self.entries_options, width=12, state="readonly")
        self.entry_combo.pack(side="left", padx=(0,8))
        ttk.Label(playbook_frame, text="Time Frame Entry:").pack(side="left", padx=(5,2))
        self.timeframe_combo = ttk.Combobox(playbook_frame, textvariable=self.timeframe_var, values=TIMEFRAME_ENTRIES, width=6, state="readonly")
        self.timeframe_combo.pack(side="left", padx=(0,8))

        # Timeframe Screenshots Frame
        tfss_frame = ttk.LabelFrame(self.journal_tab, text="Timeframe Screenshots")
        tfss_frame.pack(padx=10, pady=8, fill="x")
        self.tf_screenshots = {
            "D1": {"before": None, "after": None},
            "H4": {"before": None, "after": None},
            "H1": {"before": None, "after": None}
        }
        self.tf_img_labels = {}
        row = 0
        for tf in ["D1", "H4", "H1"]:
            ttk.Label(tfss_frame, text=tf, font=("Segoe UI", 10, "bold")).grid(row=row, column=0, padx=8, pady=6, sticky="e")
            before_lbl = ttk.Label(tfss_frame, text="No before screenshot")
            before_lbl.grid(row=row, column=1, padx=5)
            ttk.Button(tfss_frame, text="Attach Before", command=lambda t=tf: self.attach_tf_img(t, "before")).grid(row=row, column=2, padx=2)
            after_lbl = ttk.Label(tfss_frame, text="No after screenshot")
            after_lbl.grid(row=row, column=3, padx=5)
            ttk.Button(tfss_frame, text="Attach After", command=lambda t=tf: self.attach_tf_img(t, "after")).grid(row=row, column=4, padx=2)
            self.tf_img_labels[(tf, "before")] = before_lbl
            self.tf_img_labels[(tf, "after")] = after_lbl
            row += 1

        # Add Trade Button
        ttk.Button(self.journal_tab, text="Add Trade", command=self.add_trade).pack(pady=8)

        # Variable Traces for dynamic updates
        self.sl_pips_var.trace_add("write", lambda *a: self.update_sl_from_pips())
        self.sl_price_var.trace_add("write", lambda *a: self.update_sl_from_price())
        self.tp_pips_var.trace_add("write", lambda *a: self.update_tp_from_pips())
        self.tp_price_var.trace_add("write", lambda *a: self.update_tp_from_price())
        self.entry_price_var.trace_add("write", lambda *a: self.sync_all_sl_tp())
        self.lot_size_var.trace_add("write", lambda *a: self.sync_all_sl_tp())
        self.account_balance_var.trace_add("write", lambda *a: self.sync_all_sl_tp())
        self.trade_type_var.trace_add("write", lambda *a: self.sync_all_sl_tp())
        self.timezone_var.trace_add("write", self.update_market_session)
        self.trade_date_var.trace_add("write", self.update_market_session)
        self.trade_time_var.trace_add("write", self.update_market_session)
        self.update_market_session()

    def build_review_tab(self):
        review_tab_frame = ttk.Frame(self.review_tab)
        review_tab_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("id", "symbol", "trade_type", "date", "time", "entry_price", "lot_size",
                   "sl_price", "tp_price", "setup", "entry", "outcome", "pnl", "notes",
                   "acc_bal_before", "total_pips", "win_loss", "profit_usd", "gain_pct",
                   "acc_bal_after", "mdd_pips", "mdd_usd", "sl_to_be")
        self.trades_tree = ttk.Treeview(review_tab_frame, columns=columns, show="headings")

        self.trades_tree.heading("id", text="ID")
        self.trades_tree.heading("symbol", text="Symbol")
        self.trades_tree.heading("trade_type", text="Type")
        self.trades_tree.heading("date", text="Date")
        self.trades_tree.heading("time", text="Time")
        self.trades_tree.heading("entry_price", text="Entry Price")
        self.trades_tree.heading("lot_size", text="Lot Size")
        self.trades_tree.heading("sl_price", text="SL Price")
        self.trades_tree.heading("tp_price", text="TP Price")
        self.trades_tree.heading("setup", text="Setup")
        self.trades_tree.heading("entry", text="Entry")
        self.trades_tree.heading("outcome", text="Outcome")
        self.trades_tree.heading("pnl", text="P&L")
        self.trades_tree.heading("notes", text="Notes")

        self.trades_tree.heading("acc_bal_before", text="Bal Before")
        self.trades_tree.heading("total_pips", text="Total Pips")
        self.trades_tree.heading("win_loss", text="W/L")
        self.trades_tree.heading("profit_usd", text="Profit ($)")
        self.trades_tree.heading("gain_pct", text="Gain (%)")
        self.trades_tree.heading("acc_bal_after", text="Bal After")
        self.trades_tree.heading("mdd_pips", text="Max DD Pips")
        self.trades_tree.heading("mdd_usd", text="Max DD ($)")
        self.trades_tree.heading("sl_to_be", text="SL to BE?")

        self.trades_tree.column("id", width=30, anchor="center")
        self.trades_tree.column("symbol", width=70, anchor="center")
        self.trades_tree.column("trade_type", width=60, anchor="center")
        self.trades_tree.column("date", width=90, anchor="center")
        self.trades_tree.column("time", width=60, anchor="center")
        self.trades_tree.column("entry_price", width=80, anchor="center")
        self.trades_tree.column("lot_size", width=60, anchor="center")
        self.trades_tree.column("sl_price", width=80, anchor="center")
        self.trades_tree.column("tp_price", width=80, anchor="center")
        self.trades_tree.column("setup", width=90, anchor="center")
        self.trades_tree.column("entry", width=80, anchor="center")
        self.trades_tree.column("outcome", width=90, anchor="center")
        self.trades_tree.column("pnl", width=70, anchor="center")
        self.trades_tree.column("notes", width=150, anchor="w")

        self.trades_tree.column("acc_bal_before", width=90, anchor="center")
        self.trades_tree.column("total_pips", width=80, anchor="center")
        self.trades_tree.column("win_loss", width=50, anchor="center")
        self.trades_tree.column("profit_usd", width=80, anchor="center")
        self.trades_tree.column("gain_pct", width=70, anchor="center")
        self.trades_tree.column("acc_bal_after", width=90, anchor="center")
        self.trades_tree.column("mdd_pips", width=80, anchor="center")
        self.trades_tree.column("mdd_usd", width=80, anchor="center")
        self.trades_tree.column("sl_to_be", width=80, anchor="center")

        self.trades_tree.pack(fill="both", expand=True)

        h_scrollbar = ttk.Scrollbar(review_tab_frame, orient="horizontal", command=self.trades_tree.xview)
        v_scrollbar = ttk.Scrollbar(review_tab_frame, orient="vertical", command=self.trades_tree.yview)
        self.trades_tree.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)

        h_scrollbar.pack(side="bottom", fill="x")
        v_scrollbar.pack(side="right", fill="y")

        review_button_frame = ttk.Frame(review_tab_frame)
        review_button_frame.pack(pady=5)
        ttk.Button(review_button_frame, text="Review/Edit Selected Trade", command=self.open_review_popup_from_button).pack()

        self.trades_tree.bind('<Double-1>', self.open_review_popup)
        self.refresh_trades_tree()

    def build_playbook_tab(self):
        """Builds the new Playbook management tab."""
        # Create a canvas to hold the playbook frames, allowing for scrolling
        canvas = tk.Canvas(self.playbook_tab, borderwidth=0, background="#f8f8f8")
        scroll_y = ttk.Scrollbar(self.playbook_tab, orient="vertical", command=canvas.yview)
        scroll_x = ttk.Scrollbar(self.playbook_tab, orient="horizontal", command=canvas.xview) # Horizontal scrollbar
        
        canvas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set) # Configure both scrollbars
        canvas.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x") # Pack horizontal scrollbar at the bottom

        # Create a frame inside the canvas to put the content
        main_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=main_frame, anchor="nw", tags="main_frame")

        # Configure scrolling
        main_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Bind mouse wheel scrolling to the canvas for all platforms
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
        canvas.bind_all("<Button-4>", lambda event: canvas.yview_scroll(-1, "units")) # For Linux
        canvas.bind_all("<Button-5>", lambda event: canvas.yview_scroll(1, "units")) # For Linux


        # Create a dictionary to hold references to listboxes and entry widgets
        self.playbook_widgets = {}

        # Define categories and their corresponding files/option lists
        categories = {
            "Setup Options": {"options": self.setups_options, "file": SETUP_FILE},
            "Entry Options": {"options": self.entries_options, "file": ENTRY_FILE},
            "Stop Loss Reasons": {"options": self.sl_logic_options, "file": SL_REASONS_FILE},
            "Take Profit Reasons": {"options": self.tp_logic_options, "file": TP_REASONS_FILE},
            "Partial Close Reasons": {"options": self.partial_close_reasons_options, "file": PARTIAL_CLOSE_REASONS_FILE},
        }

        col_idx = 0
        row_idx = 0
        max_cols = 2 # Number of columns for side-by-side layout

        for category_name, data in categories.items():
            frame = ttk.LabelFrame(main_frame, text=category_name, padding="10")
            frame.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="nsew") # Use nsew for expansion

            # Make columns/rows in the frame expandable
            frame.grid_columnconfigure(0, weight=1)
            frame.grid_rowconfigure(0, weight=1)

            # Listbox to display options
            listbox = tk.Listbox(frame, height=8, width=35) # Adjusted height and width for smaller frames
            listbox.grid(row=0, column=0, columnspan=2, padx=(0,5), sticky="nsew") # Use grid now
            
            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
            scrollbar.grid(row=0, column=2, sticky="ns") # Place scrollbar next to listbox
            listbox.config(yscrollcommand=scrollbar.set)

            # Entry for adding/editing options
            entry_var = tk.StringVar()
            entry = ttk.Entry(frame, textvariable=entry_var, width=30) # Adjusted width
            entry.grid(row=1, column=0, columnspan=3, pady=(5,0), sticky="ew")

            # Buttons
            button_frame = ttk.Frame(frame)
            button_frame.grid(row=2, column=0, columnspan=3, pady=(5,0))

            add_btn = ttk.Button(button_frame, text="Add", command=lambda lb=listbox, ev=entry_var, opt_list=data["options"], fpath=data["file"]: self._add_playbook_item(lb, ev, opt_list, fpath))
            add_btn.pack(side="left", padx=(0,5))

            edit_btn = ttk.Button(button_frame, text="Edit Selected", command=lambda lb=listbox, ev=entry_var, opt_list=data["options"], fpath=data["file"]: self._edit_playbook_item(lb, ev, opt_list, fpath))
            edit_btn.pack(side="left", padx=(0,5))

            delete_btn = ttk.Button(button_frame, text="Delete Selected", command=lambda lb=listbox, opt_list=data["options"], fpath=data["file"]: self._delete_playbook_item(lb, opt_list, fpath))
            delete_btn.pack(side="left")

            # Store references
            self.playbook_widgets[category_name] = {
                "listbox": listbox,
                "entry_var": entry_var,
                "options_list": data["options"], # Reference to the actual list in self
                "filepath": data["file"]
            }

            # Populate listbox initially
            self._populate_listbox(listbox, data["options"])

            # Bind selection to entry field for editing
            listbox.bind('<<ListboxSelect>>', lambda e, ev=entry_var, lb=listbox: self._on_playbook_select(e, ev, lb))

            col_idx += 1
            if col_idx >= max_cols:
                col_idx = 0
                row_idx += 1
        
        # Ensure the main_frame's columns expand to fill canvas
        for i in range(max_cols):
            main_frame.grid_columnconfigure(i, weight=1)
        
        # Add an empty row at the bottom to push content up for scrolling if needed
        main_frame.grid_rowconfigure(row_idx, weight=1)


    def _populate_listbox(self, listbox, options_list):
        """Helper to populate a listbox with options."""
        listbox.delete(0, tk.END)
        for item in options_list:
            listbox.insert(tk.END, item)

    def _on_playbook_select(self, event, entry_var, listbox):
        """Event handler for listbox selection to populate entry field."""
        selection = listbox.curselection()
        if selection:
            index = selection[0]
            selected_item = listbox.get(index)
            entry_var.set(selected_item)

    def _add_playbook_item(self, listbox, entry_var, options_list, filepath):
        """Adds a new item to the playbook category."""
        new_item = entry_var.get().strip()
        if new_item and new_item not in options_list:
            options_list.append(new_item)
            options_list.sort() # Keep it sorted
            self.save_playbook_options(options_list, filepath)
            self._populate_listbox(listbox, options_list)
            entry_var.set("") # Clear entry
            self.refresh_journal_dropdowns() # Update dropdowns on main tab
        elif not new_item:
            messagebox.showwarning("Input Error", "Please enter an item to add.")
        else:
            messagebox.showinfo("Duplicate", f"'{new_item}' already exists in this category.")

    def _edit_playbook_item(self, listbox, entry_var, options_list, filepath):
        """Edits the selected item in the playbook category."""
        selection = listbox.curselection()
        if not selection:
            messagebox.showwarning("Selection Error", "Please select an item to edit.")
            return

        old_item_index = selection[0]
        old_item = listbox.get(old_item_index)
        new_item = entry_var.get().strip()

        if not new_item:
            messagebox.showwarning("Input Error", "Please enter a new value for the selected item.")
            return

        if new_item == old_item:
            messagebox.showinfo("No Change", "The new item is the same as the old one.")
            return

        if new_item in options_list and new_item != old_item: # Check if new item already exists and is different from old
            messagebox.showinfo("Duplicate", f"'{new_item}' already exists in this category.")
            return

        # Update the actual list
        options_list[options_list.index(old_item)] = new_item
        options_list.sort() # Re-sort after edit

        self.save_playbook_options(options_list, filepath)
        self._populate_listbox(listbox, options_list)
        entry_var.set("") # Clear entry
        self.refresh_journal_dropdowns() # Update dropdowns on main tab

    def _delete_playbook_item(self, listbox, options_list, filepath):
        """Deletes the selected item from the playbook category."""
        selection = listbox.curselection()
        if not selection:
            messagebox.showwarning("Selection Error", "Please select an item to delete.")
            return

        item_to_delete = listbox.get(selection[0])
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{item_to_delete}'?"):
            if item_to_delete in options_list:
                options_list.remove(item_to_delete)
                self.save_playbook_options(options_list, filepath)
                self._populate_listbox(listbox, options_list)
                self.refresh_journal_dropdowns() # Update dropdowns on main tab

    def attach_tf_img(self, tf, when):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")])
        if path:
            self.tf_screenshots[tf][when] = path
            lbl = self.tf_img_labels[(tf, when)]
            self._show_img_thumbnail(path, lbl)
        else:
            lbl = self.tf_img_labels[(tf, when)]
            lbl["text"] = f"No {when} screenshot"

    def add_trade(self):
        symbol = self.symbol_var.get()
        tf = self.timeframe_var.get() or "1h"

        info = {
            "symbol": self.symbol_var.get(),
            "timeframe": self.timeframe_var.get() or "1h",
            "trade_type": self.trade_type_var.get(),
            "trade_date": self.trade_date_var.get(),
            "trade_time": self.trade_time_var.get(),
            "market_session": self.market_session_var.get(),
            "timezone": self.timezone_var.get(),
            "entry_price": self.entry_price_var.get(),
            "lot_size": self.lot_size_var.get(),
            "sl_pips": self.sl_pips_var.get(),
            "sl_price": self.sl_price_var.get(),
            "sl_loss": self.sl_loss_var.get(),
            "sl_loss_pct": self.sl_loss_pct_var.get(),
            "sl_reason": self.sl_reason_var.get(),
            "tp_pips": self.tp_pips_var.get(),
            "tp_price": self.tp_price_var.get(),
            "tp_profit": self.tp_profit_var.get(),
            "tp_profit_pct": self.tp_profit_pct_var.get(),
            "tp_reason": self.tp_reason_var.get(),
            "setup": self.setup_var.get(),
            "entry": self.entry_var.get(),
            "account_balance": self.account_balance_var.get(),
        }

        tf_shots = {k: v.copy() for k, v in self.tf_screenshots.items()}

        trade = Trade(symbol, tf, info=info, tf_screenshots=tf_shots, sl_to_be=False)
        self.trades.append(trade)
        self.save_trades()
        self.refresh_trades_tree()
        self.update_stats_bar()

        self.symbol_var.set("XAUUSD")
        self.timeframe_var.set("")
        self.trade_type_var.set("Buy")
        self._update_trade_type_buttons()
        self.trade_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        self.trade_time_var.set(datetime.now().strftime("%H:%M"))
        self.entry_price_var.set(0.0)
        self.lot_size_var.set(0.0)
        self.sl_pips_var.set(0.0)
        self.sl_price_var.set(0.0)
        self.sl_reason_var.set("")
        self.tp_pips_var.set(0.0)
        self.tp_price_var.set(0.0)
        self.tp_reason_var.set("")
        self.setup_var.set("")
        self.entry_var.set("")
        self.sync_all_sl_tp()

        for tf in ["D1", "H4", "H1"]:
            for when in ["before", "after"]:
                self.tf_screenshots[tf][when] = None
                lbl = self.tf_img_labels[(tf, when)]
                lbl["text"] = f"No {when} screenshot"
                lbl["image"] = ""

    def save_trades(self):
        try:
            data_to_save = {
                "trades": [t.to_dict() for t in self.trades],
                "account_balance": self.account_balance_var.get()
            }
            with open(TRADES_FILE, "w", encoding="utf-8") as f:
                json.dump(data_to_save, f, indent=2)
        except Exception as e:
            messagebox.showerror("Save Failed", str(e))

    def load_trades(self):
        if not os.path.exists(TRADES_FILE):
            self.trades = []
            self.account_balance_var.set(10000.00)
            self.refresh_trades_tree()
            self.update_stats_bar()
            return
        try:
            with open(TRADES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.trades = [Trade.from_dict(d) for d in data.get("trades", [])]
            self.account_balance_var.set(data.get("account_balance", 10000.00))

            self.refresh_trades_tree()
            self.update_stats_bar()
        except json.JSONDecodeError as e:
            response = messagebox.askyesno(
                "Corrupted Journal File",
                f"Error decoding journal file: {e}\n\n"
                "The journal file appears to be corrupted. Do you want to try to back up the existing file "
                "and start a new, empty journal? (Selecting 'No' will exit the application.)"
            )
            if response:
                if os.path.exists(TRADES_FILE):
                    try:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        os.rename(TRADES_FILE, f"{TRADES_FILE}.bak_{timestamp}")
                        messagebox.showinfo("Backup Created", f"Corrupted file backed up as '{TRADES_FILE}.bak_{timestamp}'")
                    except Exception as backup_e:
                        messagebox.showwarning("Backup Failed", f"Could not back up corrupted file: {backup_e}\nStarting new journal anyway.")
                self.trades = []
                self.account_balance_var.set(10000.00)
                self.refresh_trades_tree()
                self.update_stats_bar()
                self.save_trades()
            else:
                self.quit()
        except Exception as e:
            messagebox.showerror("Load Failed", f"An unexpected error occurred during load: {e}")

    def refresh_trades_tree(self):
        for i in self.trades_tree.get_children():
            self.trades_tree.delete(i)
        for idx, trade in enumerate(self.trades):
            total_pnl_for_display = 0.0
            total_pips_for_display = 0.0
            win_loss_status = "Open"

            initial_balance = trade.info.get("account_balance", 0.0)
            original_lot_size = float(trade.info.get("lot_size", 0.0))
            entry_price = float(trade.info.get("entry_price", 0.0))
            trade_type = trade.info.get("trade_type", "Buy")

            for pc in trade.partial_closes:
                total_pnl_for_display += pc.get("pnl", 0.0)
                total_pips_for_display += pc.get("pips", 0.0)

            if trade.review.get("outcome") in ["Take Profit Hit", "Stoploss Hit", "Breakeven", "Other"]:
                try:
                    close_price = float(trade.review.get("price", 0))

                    closed_lot_size_partials = sum(pc.get("amount", 0) for pc in trade.partial_closes)
                    remaining_lot_size = original_lot_size - closed_lot_size_partials

                    if remaining_lot_size > 0:
                        if trade_type == "Buy":
                            pips_moved = (close_price - entry_price) / PIP_VALUE_XAUUSD
                        else:
                            pips_moved = (entry_price - close_price) / PIP_VALUE_XAUUSD

                        total_pips_for_display += pips_moved
                        total_pnl_for_display += pips_moved * remaining_lot_size * USD_PER_PIP_PER_LOT

                    if total_pnl_for_display > 0:
                        win_loss_status = "Win"
                    elif total_pnl_for_display < 0:
                        win_loss_status = "Loss"
                    else:
                        win_loss_status = "Breakeven"

                except ValueError:
                    pass
                except Exception:
                    pass

            pnl_value_main_column = f"${total_pnl_for_display:,.2f}"
            if trade.partial_closes:
                pnl_value_main_column += "*"

            profit_usd_value = f"${total_pnl_for_display:,.2f}"
            gain_pct_value = ""
            if initial_balance > 0:
                gain_pct = (total_pnl_for_display / initial_balance) * 100
                gain_pct_value = f"{gain_pct:,.2f}%"

            acc_bal_after_value = f"${initial_balance + total_pnl_for_display:,.2f}"

            mdd_pips_val = trade.review.get("max_drawdown_pips", "")
            mdd_usd_val = ""
            try:
                mdd_pips_float = float(mdd_pips_val)
                lot_size_float = float(trade.info.get("lot_size", 0.0))
                mdd_usd = mdd_pips_float * lot_size_float * USD_PER_PIP_PER_LOT
                mdd_usd_val = f"${mdd_usd:,.2f}"
            except (ValueError, TypeError):
                pass

            sl_to_be_display = "Yes" if trade.sl_to_be else "No"

            vals = [
                idx + 1,
                trade.symbol,
                trade.info.get("trade_type", ""),
                trade.info.get("trade_date", ""),
                trade.info.get("trade_time", ""),
                trade.info.get("entry_price", ""),
                trade.info.get("lot_size", ""),
                trade.info.get("sl_price", ""),
                trade.info.get("tp_price", ""),
                trade.info.get("setup", ""),
                trade.info.get("entry", ""),
                trade.review.get("outcome", ""),
                pnl_value_main_column,
                trade.review.get("notes", ""),
                f"${initial_balance:,.2f}",
                f"{total_pips_for_display:,.1f}",
                win_loss_status,
                profit_usd_value,
                gain_pct_value,
                acc_bal_after_value,
                f"{mdd_pips_val}",
                mdd_usd_val,
                sl_to_be_display
            ]
            self.trades_tree.insert("", "end", iid=idx, values=vals)

    def update_stats_bar(self):
        total = len(self.trades)
        total_realized_pnl = 0.0
        for trade in self.trades:
            for pc in trade.partial_closes:
                total_realized_pnl += pc.get("pnl", 0.0)

            if trade.review.get("outcome") in ["Take Profit Hit", "Stoploss Hit", "Breakeven", "Other"]:
                try:
                    entry_price = float(trade.info.get("entry_price", 0))
                    close_price = float(trade.review.get("price", 0))
                    original_lot_size = float(trade.info.get("lot_size", 0))
                    trade_type = trade.info.get("trade_type", "Buy")

                    closed_lot_size_partials = sum(pc.get("amount", 0) for pc in trade.partial_closes)
                    remaining_lot_size = original_lot_size - closed_lot_size_partials

                    if remaining_lot_size > 0:
                        if trade_type == "Buy":
                            pips_moved = (close_price - entry_price) / PIP_VALUE_XAUUSD
                        else:
                            pips_moved = (entry_price - close_price) / PIP_VALUE_XAUUSD
                        total_realized_pnl += pips_moved * remaining_lot_size * USD_PER_PIP_PER_LOT
                except ValueError:
                    pass

        wins = sum(1 for t in self.trades if t.review.get("outcome", "").lower() == "take profit hit")
        losses = sum(1 for t in self.trades if t.review.get("outcome", "").lower() == "stoploss hit")
        wr = int(round(100 * wins / total)) if total > 0 else 0

        self.trade_stats_var.set(f"Trade Count {total}  Win {wins}  Loss {losses}  {wr}%WR | Realized P&L: ${total_realized_pnl:,.2f}")

    def open_review_popup_from_button(self):
        selected = self.trades_tree.focus()
        if not selected:
            messagebox.showinfo("No Trade Selected", "Please select a trade from the list to review or edit.")
            return
        self.open_review_popup(None)

    def open_review_popup(self, event=None):
        selected = self.trades_tree.focus()
        if not selected:
            return
        idx = int(selected)
        trade = self.trades[idx]

        popup = tk.Toplevel(self)
        popup.title(f"Edit/Review Trade ({trade.symbol} {trade.timeframe})")
        popup.geometry("850x850")
        popup.grab_set()

        canvas = tk.Canvas(popup, borderwidth=0, background="#f8f8f8")
        frame = ttk.Frame(canvas)
        vsb = tk.Scrollbar(popup, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=frame, anchor="nw")

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame.bind("<Configure>", on_frame_configure)
        
        # Bind mouse wheel scrolling to the canvas for all platforms in popup
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
        canvas.bind_all("<Button-4>", lambda event: canvas.yview_scroll(-1, "units")) # For Linux
        canvas.bind_all("<Button-5>", lambda event: canvas.yview_scroll(1, "units")) # For Linux


        info = trade.info
        review = trade.review

        info_frame = ttk.LabelFrame(frame, text="Trade Details (View Only)")
        info_frame.pack(padx=10, pady=5, fill="x")

        def viewrow(parent, label, val, row, col=0, colspan=1):
            ttk.Label(parent, text=label, font=("Segoe UI", 9, "bold")).grid(row=row, column=col, sticky="w", padx=4, pady=2)
            ttk.Label(parent, text=val, font=("Segoe UI", 9)).grid(row=row, column=col+1, sticky="w", padx=4, pady=2, columnspan=colspan)

        viewrow(info_frame, "Symbol:", info.get("symbol", trade.symbol), 0)
        viewrow(info_frame, "Timeframe:", info.get("timeframe", trade.timeframe), 0, col=2)
        viewrow(info_frame, "Trade Type:", info.get("trade_type", ""), 1)
        viewrow(info_frame, "Trade Date:", info.get("trade_date", ""), 1, col=2)
        viewrow(info_frame, "Trade Time:", info.get("trade_time", ""), 2)
        viewrow(info_frame, "Setup:", info.get("setup", ""), 2, col=2)
        viewrow(info_frame, "Entry Type:", info.get("entry", ""), 3)
        viewrow(info_frame, "Entry Price:", info.get("entry_price", ""), 3, col=2)
        viewrow(info_frame, "Lot Size (Initial):", info.get("lot_size", ""), 4)
        viewrow(info_frame, "Market Session:", info.get("market_session", ""), 4, col=2)
        viewrow(info_frame, "Timezone:", info.get("timezone", ""), 5)
        viewrow(info_frame, "Account Balance (Entry):", f"${info.get('account_balance', 0):,.2f}", 5, col=2)

        ttk.Separator(info_frame, orient="horizontal").grid(row=6, column=0, columnspan=4, sticky="ew", pady=6)

        viewrow(info_frame, "Stop Loss (Price):", info.get("sl_price", ""), 7, col=0)
        viewrow(info_frame, "Take Profit (Price):", info.get("tp_price", ""), 7, col=2)
        viewrow(info_frame, "SL Pips:", info.get("sl_pips", ""), 8, col=0)
        viewrow(info_frame, "TP Pips:", info.get("tp_pips", ""), 8, col=2)
        viewrow(info_frame, "SL Reason:", info.get("sl_reason", ""), 9, col=0)
        viewrow(info_frame, "TP Reason:", info.get("tp_reason", ""), 9, col=2)

        partial_frame = ttk.LabelFrame(frame, text="Partial Closures")
        partial_frame.pack(padx=10, pady=5, fill="x")

        partial_columns = ("timestamp", "amount", "price", "pips", "pnl", "reason_for_close")
        partial_tree = ttk.Treeview(partial_frame, columns=partial_columns, show="headings", height=5)
        partial_tree.heading("timestamp", text="Timestamp")
        partial_tree.heading("amount", text="Amount (Lots)")
        partial_tree.heading("price", text="Price")
        partial_tree.heading("pips", text="Pips")
        partial_tree.heading("pnl", text="P&L")
        partial_tree.heading("reason_for_close", text="Reason for Close")

        partial_tree.column("timestamp", width=120, anchor="center")
        partial_tree.column("amount", width=80, anchor="center")
        partial_tree.column("price", width=80, anchor="center")
        partial_tree.column("pips", width=60, anchor="center")
        partial_tree.column("pnl", width=80, anchor="center")
        partial_tree.column("reason_for_close", width=200, anchor="w")

        partial_tree.pack(fill="x", padx=5, pady=5)

        partial_v_scrollbar = ttk.Scrollbar(partial_frame, orient="vertical", command=partial_tree.yview)
        partial_tree.configure(yscrollcommand=partial_v_scrollbar.set)
        partial_v_scrollbar.pack(side="right", fill="y")

        add_partial_frame = ttk.Frame(partial_frame)
        add_partial_frame.pack(fill="x", pady=5, padx=5)

        ttk.Label(add_partial_frame, text="Amount (Lots):").grid(row=0, column=0, padx=2, pady=2)
        partial_amount_var = tk.DoubleVar(value=0.0)
        partial_amount_entry = ttk.Entry(add_partial_frame, textvariable=partial_amount_var, width=8)
        partial_amount_entry.grid(row=0, column=1, padx=2, pady=2)
        self._on_focus_select_all(partial_amount_entry, partial_amount_var)

        ttk.Label(add_partial_frame, text="Pips:").grid(row=0, column=2, padx=2, pady=2)
        partial_pips_var = tk.DoubleVar(value=0.0)
        partial_pips_entry = ttk.Entry(add_partial_frame, textvariable=partial_pips_var, width=8)
        partial_pips_entry.grid(row=0, column=3, padx=2, pady=2)
        self._on_focus_select_all(partial_pips_entry, partial_pips_var)

        ttk.Label(add_partial_frame, text="Price:").grid(row=0, column=4, padx=2, pady=2)
        partial_price_var = tk.DoubleVar(value=0.0)
        partial_price_entry = ttk.Entry(add_partial_frame, textvariable=partial_price_var, width=8, state="normal")
        partial_price_entry.grid(row=0, column=5, padx=2, pady=2)
        self._on_focus_select_all(partial_price_entry, partial_price_var)

        ttk.Label(add_partial_frame, text="Reason for Close:").grid(row=0, column=6, padx=2, pady=2)
        partial_reason_var = tk.StringVar(value="")
        # Updated to use self.partial_close_reasons_options
        partial_reason_combo = ttk.Combobox(add_partial_frame, textvariable=partial_reason_var,
                                            values=self.partial_close_reasons_options, width=25, state="readonly")
        partial_reason_combo.grid(row=0, column=7, padx=2, pady=2)

        committed_closed_lots = sum(pc.get("amount", 0.0) for pc in trade.partial_closes)
        original_lot_size = info.get("lot_size", 0.0)
        entry_price = float(info.get("entry_price", 0))
        trade_type = info.get("trade_type", "Buy")

        remaining_lot_preview_var = tk.DoubleVar(value=original_lot_size - committed_closed_lots)
        current_partial_pnl_display_var = tk.StringVar(value="$0.00")

        ttk.Label(add_partial_frame, text="Remaining Lots:").grid(row=1, column=0, padx=2, pady=2, sticky="w")
        ttk.Label(add_partial_frame, textvariable=remaining_lot_preview_var, font=("Segoe UI", 9, "bold")).grid(row=1, column=1, padx=2, pady=2, sticky="w")

        ttk.Label(add_partial_frame, text="Current P&L:").grid(row=1, column=2, padx=2, pady=2, sticky="w")
        ttk.Label(add_partial_frame, textvariable=current_partial_pnl_display_var, font=("Segoe UI", 9, "bold")).grid(row=1, column=3, padx=2, pady=2, sticky="w")

        is_updating_partial_sync = False

        def _update_partial_pips_from_price(*args):
            nonlocal is_updating_partial_sync
            if is_updating_partial_sync: return
            is_updating_partial_sync = True
            try:
                price_val_str = partial_price_var.get()
                price = float(price_val_str)

                if price <= 0:
                    new_pips = 0.0
                else:
                    if trade_type == "Buy":
                        new_pips = round((price - entry_price) / PIP_VALUE_XAUUSD, 1)
                    else:
                        new_pips = round((entry_price - price) / PIP_VALUE_XAUUSD, 1)

                if abs(partial_pips_var.get() - new_pips) > 1e-6:
                    partial_pips_var.set(new_pips)
            except (ValueError, tk.TclError):
                pass
            finally:
                is_updating_partial_sync = False
                _update_partial_previews()

        def _update_partial_price_from_pips(*args):
            nonlocal is_updating_partial_sync
            if is_updating_partial_sync: return
            is_updating_partial_sync = True
            try:
                pips_val_str = partial_pips_var.get()
                pips = float(pips_val_str)

                if trade_type == "Buy":
                    new_price = entry_price + (pips * PIP_VALUE_XAUUSD)
                else:
                    new_price = entry_price - (pips * PIP_VALUE_XAUUSD)

                if abs(partial_price_var.get() - new_price) > 1e-6:
                    partial_price_var.set(round(new_price, 2))
            except (ValueError, tk.TclError):
                pass
            finally:
                is_updating_partial_sync = False
                _update_partial_previews()

        def _update_partial_previews(*args):
            try:
                current_partial_amount = partial_amount_var.get()
                try:
                    current_partial_pips = partial_pips_var.get()
                except tk.TclError:
                    current_partial_pips = 0.0

                new_remaining_lots = original_lot_size - committed_closed_lots - current_partial_amount
                remaining_lot_preview_var.set(max(0, round(new_remaining_lots, 2)))

                if current_partial_amount > 0:
                    calculated_pnl = current_partial_pips * current_partial_amount * USD_PER_PIP_PER_LOT
                    current_partial_pnl_display_var.set(f"${calculated_pnl:,.2f}")
                else:
                    current_partial_pnl_display_var.set("$0.00")
            except (ValueError, tk.TclError):
                remaining_lot_preview_var.set(original_lot_size - committed_closed_lots)
                current_partial_pnl_display_var.set("Invalid Input")
            except Exception:
                remaining_lot_preview_var.set(original_lot_size - committed_closed_lots)
                current_partial_pnl_display_var.set("Error")

        partial_amount_var.trace_add("write", _update_partial_previews)
        partial_pips_var.trace_add("write", _update_partial_price_from_pips)
        partial_price_var.trace_add("write", _update_partial_pips_from_price)

        _update_partial_previews()

        def add_partial_close_entry():
            try:
                amount = partial_amount_var.get()
                price = partial_price_var.get()
                pips = partial_pips_var.get()
                reason_for_close = partial_reason_var.get().strip()

                if amount <= 0 or (price == 0 and pips == 0):
                    messagebox.showwarning("Invalid Input", "Amount must be positive and Price/Pips must not be zero for partial close.")
                    return

                current_remaining_lots_before_this_partial = original_lot_size - sum(pc.get("amount", 0.0) for pc in trade.partial_closes)
                if amount > current_remaining_lots_before_this_partial + 0.0001:
                    messagebox.showwarning("Invalid Input", f"Partial amount ({amount}) exceeds remaining lot size ({current_remaining_lots_before_this_partial:.2f}).")
                    return

                if not reason_for_close:
                     messagebox.showwarning("Missing Information", "Please select a reason for the partial close.")
                     return

                partial_pnl = pips * amount * USD_PER_PIP_PER_LOT

                new_partial = {
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "amount": amount,
                    "price": price,
                    "pips": pips,
                    "reason_for_close": reason_for_close,
                    "pnl": partial_pnl
                }
                trade.partial_closes.append(new_partial)

                nonlocal committed_closed_lots
                committed_closed_lots = sum(pc.get("amount", 0.0) for pc in trade.partial_closes)

                self.refresh_partial_tree(partial_tree, trade)
                self.save_trades()
                self.refresh_trades_tree()
                self.update_stats_bar()

                partial_amount_var.set(0.0)
                partial_pips_var.set(0.0)
                partial_price_var.set(0.0)
                partial_reason_var.set("")
                _update_partial_previews()

                messagebox.showinfo("Partial Close Added", "Partial close recorded successfully.")

            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter valid numbers for amount, pips/price.")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

        ttk.Button(add_partial_frame, text="Add Partial Close", command=add_partial_close_entry).grid(row=1, column=6, columnspan=2, pady=5)

        self.refresh_partial_tree(partial_tree, trade)

        review_frame = ttk.LabelFrame(frame, text="Final Review & Outcome")
        review_frame.pack(padx=10, pady=5, fill="x")

        outcome_var = tk.StringVar(value=review.get("outcome", ""))
        close_price_var = tk.StringVar(value=review.get("price", ""))

        exit_time_var = tk.StringVar(value=review.get("exit_time", ""))
        time_in_trade_var = tk.StringVar(value="")

        max_drawdown_pips_var = tk.StringVar(value=review.get("max_drawdown_pips", ""))
        max_drawdown_usd_var = tk.StringVar(value="")
        max_drawdown_price_var = tk.StringVar(value="")

        sl_to_be_var = tk.BooleanVar(value=trade.sl_to_be)

        def sync_price_field(*args):
            outcome = outcome_var.get()
            try:
                entry_price = float(info.get("entry_price", 0))
                tp_price = float(info.get("tp_price", 0))
                sl_price = float(info.get("sl_price", 0))
            except ValueError:
                entry_price = tp_price = sl_price = 0.0

            if outcome == "Take Profit Hit":
                close_price_var.set(f"{tp_price:.2f}")
                price_entry.config(state="readonly")
            elif outcome == "Stoploss Hit":
                close_price_var.set(f"{sl_price:.2f}")
                price_entry.config(state="readonly")
            elif outcome == "Breakeven":
                close_price_var.set(f"{entry_price:.2f}")
                price_entry.config(state="readonly")
            elif outcome == "Other":
                price_entry.config(state="normal")
                if close_price_var.get() == "0.0" or close_price_var.get() == "":
                    close_price_var.set("")
            else:
                if close_price_var.get() == "0.0" or close_price_var.get() == "":
                    close_price_var.set("")
                price_entry.config(state="readonly")

        _mdd_is_updating = False

        def _update_mdd_usd_from_pips(*args):
            nonlocal _mdd_is_updating
            if _mdd_is_updating: return
            _mdd_is_updating = True
            try:
                mdd_pips_str = max_drawdown_pips_var.get()
                mdd_pips = float(mdd_pips_str) if mdd_pips_str else 0.0

                lot_size = float(info.get("lot_size", 0.0))
                mdd_usd = mdd_pips * lot_size * USD_PER_PIP_PER_LOT
                max_drawdown_usd_var.set(f"${mdd_usd:,.2f}")

                _update_mdd_price_from_pips_internal()
            except (ValueError, tk.TclError):
                max_drawdown_usd_var.set("")
                max_drawdown_price_var.set("")
            finally:
                _mdd_is_updating = False

        def _update_mdd_pips_from_price(*args):
            nonlocal _mdd_is_updating
            if _mdd_is_updating: return
            _mdd_is_updating = True
            try:
                entry_price = float(info.get("entry_price", 0.0))
                mdd_price_str = max_drawdown_price_var.get()
                mdd_price = float(mdd_price_str) if mdd_price_str else 0.0
                trade_type = info.get("trade_type", "Buy")

                if entry_price == 0:
                    max_drawdown_pips_var.set("")
                    max_drawdown_usd_var.set("")
                    return

                if trade_type == "Buy":
                    pips = (entry_price - mdd_price) / PIP_VALUE_XAUUSD
                else:
                    pips = (mdd_price - entry_price) / PIP_VALUE_XAUUSD

                if pips < 0: pips = 0.0

                if abs(float(max_drawdown_pips_var.get() or 0.0) - pips) > 1e-6:
                    max_drawdown_pips_var.set(f"{pips:.1f}")

                _update_mdd_usd_from_pips()
            except (ValueError, tk.TclError):
                max_drawdown_pips_var.set("")
                max_drawdown_usd_var.set("")
            finally:
                _mdd_is_updating = False

        def _update_mdd_price_from_pips_internal():
            try:
                entry_price = float(info.get("entry_price", 0.0))
                mdd_pips_str = max_drawdown_pips_var.get()
                mdd_pips = float(mdd_pips_str) if mdd_pips_str else 0.0
                trade_type = info.get("trade_type", "Buy")

                if trade_type == "Buy":
                    mdd_price = entry_price - (mdd_pips * PIP_VALUE_XAUUSD)
                else:
                    mdd_price = entry_price + (mdd_pips * PIP_VALUE_XAUUSD)

                if abs(float(max_drawdown_price_var.get() or 0.0) - mdd_price) > 1e-6:
                    max_drawdown_price_var.set(f"{mdd_price:.2f}")
            except (ValueError, tk.TclError):
                max_drawdown_price_var.set("")

        outcome_options = ["", "Take Profit Hit", "Stoploss Hit", "Breakeven", "Other"]
        ttk.Label(review_frame, text="Outcome:").grid(row=0, column=0, padx=4, pady=4, sticky="e")
        outcome_combo = ttk.Combobox(
            review_frame, textvariable=outcome_var,
            values=outcome_options, width=18, state="readonly"
        )
        outcome_combo.grid(row=0, column=1, padx=4, pady=4, sticky="w")

        ttk.Label(review_frame, text="Final Close Price:").grid(row=0, column=2, padx=4, pady=4, sticky="e")
        price_entry = ttk.Entry(review_frame, textvariable=close_price_var, width=16)
        price_entry.grid(row=0, column=3, padx=4, pady=4, sticky="w")
        self._on_focus_select_all(price_entry, close_price_var)

        ttk.Label(review_frame, text="Exit Time:").grid(row=1, column=0, padx=4, pady=4, sticky="e")
        exit_time_entry = ttk.Entry(review_frame, textvariable=exit_time_var, width=16)
        exit_time_entry.grid(row=1, column=1, padx=4, pady=4, sticky="w")
        self._on_focus_select_all(exit_time_entry, exit_time_var)

        ttk.Label(review_frame, text="Time in Trade:").grid(row=1, column=2, padx=4, pady=4, sticky="e")
        time_in_trade_label = ttk.Label(review_frame, textvariable=time_in_trade_var, font=("Segoe UI", 9, "bold"))
        time_in_trade_label.grid(row=1, column=3, padx=4, pady=4, sticky="w")

        ttk.Label(review_frame, text="Max Drawdown (Pips):").grid(row=2, column=0, padx=4, pady=4, sticky="e")
        max_drawdown_pips_entry = ttk.Entry(review_frame, textvariable=max_drawdown_pips_var, width=16)
        max_drawdown_pips_entry.grid(row=2, column=1, padx=4, pady=4, sticky="w")
        self._on_focus_select_all(max_drawdown_pips_entry, max_drawdown_pips_var)

        ttk.Label(review_frame, text="Max Drawdown ($):").grid(row=2, column=2, padx=4, pady=4, sticky="e")
        max_drawdown_usd_label = ttk.Label(review_frame, textvariable=max_drawdown_usd_var, font=("Segoe UI", 9, "bold"))
        max_drawdown_usd_label.grid(row=2, column=3, padx=4, pady=4, sticky="w")

        ttk.Label(review_frame, text="Max Drawdown Price:").grid(row=3, column=0, padx=4, pady=4, sticky="e")
        max_drawdown_price_entry = ttk.Entry(review_frame, textvariable=max_drawdown_price_var, width=16)
        max_drawdown_price_entry.grid(row=3, column=1, padx=4, pady=4, sticky="w")
        self._on_focus_select_all(max_drawdown_price_entry, max_drawdown_price_var)

        sl_to_be_checkbox = ttk.Checkbutton(
            review_frame,
            text="Stop Loss Moved to Break-Even?",
            variable=sl_to_be_var
        )
        sl_to_be_checkbox.grid(row=5, column=0, columnspan=4, padx=4, pady=8, sticky="w")

        ttk.Label(review_frame, text="Final Notes:").grid(row=6, column=0, sticky="nw", padx=4, pady=4)
        notes_text = tk.Text(review_frame, width=65, height=4)
        notes_text.insert("1.0", review.get("notes", ""))
        notes_text.grid(row=6, column=1, sticky="w", padx=4, pady=4, columnspan=3)

        def calculate_time_in_trade(*args):
            try:
                trade_date = info.get("trade_date", "")
                trade_time = info.get("trade_time", "")
                exit_time = exit_time_var.get()

                if not trade_date or not trade_time or not exit_time:
                    time_in_trade_var.set("N/A")
                    return

                start_dt_str = f"{trade_date} {trade_time}"
                end_dt_str = f"{trade_date} {exit_time}"

                start_dt = datetime.strptime(start_dt_str, "%Y-%m-%d %H:%M")
                end_dt = datetime.strptime(end_dt_str, "%Y-%m-%d %H:%M")

                duration = end_dt - start_dt

                if duration < timedelta(0):
                    duration += timedelta(days=1)

                total_seconds = duration.total_seconds()
                hours = int(total_seconds // 3600)
                minutes = int((total_seconds % 3600) // 60)

                time_in_trade_var.set(f"{hours}h {minutes}m")
            except ValueError:
                time_in_trade_var.set("Invalid Time")
            except Exception:
                time_in_trade_var.set("N/A")

        outcome_var.trace_add("write", sync_price_field)
        exit_time_var.trace_add("write", calculate_time_in_trade)
        max_drawdown_pips_var.trace_add("write", _update_mdd_usd_from_pips)
        max_drawdown_price_var.trace_add("write", _update_mdd_pips_from_price)

        sync_price_field()
        calculate_time_in_trade()
        _update_mdd_usd_from_pips()

        def show_image_popup(img_path):
            try:
                img = Image.open(img_path)
                popup_img = tk.Toplevel(popup)
                popup_img.title("Screenshot Preview")
                max_w, max_h = 1200, 900
                w, h = img.size
                scale = min(1, max_w / w, max_h / h)
                if scale < 1:
                    img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
                tkimg = ImageTk.PhotoImage(img)
                lbl = tk.Label(popup_img, image=tkimg)
                lbl.image = tkimg
                lbl.pack()
                popup_img.transient(popup)
                popup_img.grab_set()
                popup_img.focus_force()
            except Exception as e:
                messagebox.showerror("Image Error", f"Could not open image:\n{e}")

        tf_frame = ttk.LabelFrame(frame, text="Timeframe Screenshots (Review/Edit)")
        tf_frame.pack(padx=10, pady=5, fill="x")

        tf_popup_img_labels = {}
        for i, tf in enumerate(["D1", "H4", "H1"]):
            ttk.Label(tf_frame, text=tf, font=("Segoe UI", 10, "bold")).grid(row=i, column=0, padx=8, pady=6, sticky="e")

            before_lbl = ttk.Label(tf_frame, text="No before screenshot", cursor="hand2", foreground="blue")
            before_lbl.grid(row=i, column=1, padx=5)
            img_path_before = trade.tf_screenshots[tf]["before"]
            if img_path_before and os.path.exists(img_path_before):
                self._show_img_thumbnail(img_path_before, before_lbl, size=(110,70))
                before_lbl.bind("<Button-1>", lambda e, path=img_path_before: show_image_popup(path))
            else:
                before_lbl["text"] = "No before screenshot"
                before_lbl.unbind("<Button-1>")
            ttk.Button(tf_frame, text="Attach/Edit Before", command=lambda t=tf: self.edit_tf_img(trade, t, "before", before_lbl)).grid(row=i, column=2, padx=2)

            after_lbl = ttk.Label(tf_frame, text="No after screenshot", cursor="hand2", foreground="blue")
            after_lbl.grid(row=i, column=3, padx=5)
            img_path_after = trade.tf_screenshots[tf]["after"]
            if img_path_after and os.path.exists(img_path_after):
                self._show_img_thumbnail(img_path_after, after_lbl, size=(110,70))
                after_lbl.bind("<Button-1>", lambda e, path=img_path_after: show_image_popup(path))
            else:
                after_lbl["text"] = "No after screenshot"
                after_lbl.unbind("<Button-1>")
            ttk.Button(tf_frame, text="Attach/Edit After", command=lambda t=tf: self.edit_tf_img(trade, t, "after", after_lbl)).grid(row=i, column=4, padx=2)

            tf_popup_img_labels[(tf, "before")] = before_lbl
            tf_popup_img_labels[(tf, "after")] = after_lbl

        def close_trade_update_balance():
            review["outcome"] = outcome_var.get()
            review["price"] = close_price_var.get()
            review["notes"] = notes_text.get("1.0", "end").strip()
            review["exit_time"] = exit_time_var.get()
            review["max_drawdown_pips"] = max_drawdown_pips_var.get()

            trade.sl_to_be = sl_to_be_var.get()

            final_pnl = 0.0

            for pc in trade.partial_closes:
                final_pnl += pc.get("pnl", 0.0)

            try:
                entry_price = float(info.get("entry_price", 0))
                close_price = float(review.get("price", 0))
                original_lot_size = float(info.get("lot_size", 0))
                trade_type = info.get("trade_type", "Buy")

                closed_lot_size_partials = sum(pc.get("amount", 0) for pc in trade.partial_closes)
                remaining_lot_size = original_lot_size - closed_lot_size_partials

                if remaining_lot_size > 0:
                    if trade_type == "Buy":
                        pips_moved = (close_price - entry_price) / PIP_VALUE_XAUUSD
                    else:
                        pips_moved = (entry_price - close_price) / PIP_VALUE_XAUUSD

                    final_pnl += pips_moved * remaining_lot_size * USD_PER_PIP_PER_LOT
            except ValueError:
                messagebox.showwarning("Calculation Warning", "Could not calculate final P&L. Ensure entry and close prices are valid numbers.")
            except Exception as e:
                messagebox.showwarning("Calculation Error", f"An error occurred during final P&L calculation: {e}")

            self.account_balance_var.set(self.account_balance_var.get() + final_pnl)

            self.save_trades()
            self.refresh_trades_tree()
            self.update_stats_bar()
            popup.destroy()

        ttk.Button(frame, text="Close Trade (Save & Update Balance)", command=close_trade_update_balance).pack(pady=12)
        ttk.Button(frame, text="Cancel", command=popup.destroy).pack(pady=2)

    def refresh_partial_tree(self, treeview, trade):
        for i in treeview.get_children():
            treeview.delete(i)

        for pc_idx, pc in enumerate(trade.partial_closes):
            partial_pnl_formatted = f"${pc.get('pnl', 0.0):,.2f}"
            treeview.insert(
                "", "end",
                values=(
                    pc.get("timestamp", ""),
                    pc.get("amount", 0.0),
                    pc.get("price", 0.0),
                    pc.get("pips", 0.0),
                    partial_pnl_formatted,
                    pc.get("reason_for_close", "")
                )
            )

    def edit_tf_img(self, trade, tf, when, lbl):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")])
        if path:
            trade.tf_screenshots[tf][when] = path
            self._show_img_thumbnail(path, lbl, size=(110,70))
            lbl.bind("<Button-1>", lambda e, path=path: self.show_image_popup_from_review(path))
            self.save_trades()
        else:
            lbl["text"] = f"No {when} screenshot"
            lbl["image"] = ""
            lbl.unbind("<Button-1>")

    def _show_img_thumbnail(self, path, label, size=(80, 50)):
        try:
            img = Image.open(path)
            img.thumbnail(size)
            img = ImageTk.PhotoImage(img)
            label.image = img
            label.configure(image=img, text="")
        except Exception:
            label.configure(text="(missing)", image="")

    def show_image_popup_from_review(self, img_path):
        try:
            img = Image.Image.open(img_path)
            popup_img = tk.Toplevel()
            popup_img.title("Screenshot Preview")
            max_w, max_h = 1200, 900
            w, h = img.size
            scale = min(1, max_w / w, max_h / h)
            if scale < 1:
                img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
            tkimg = ImageTk.PhotoImage(img)
            lbl = tk.Label(popup_img, image=tkimg)
            lbl.image = tkimg
            lbl.pack()
            popup_img.focus_force()
        except Exception as e:
            messagebox.showerror("Image Error", f"Could not open image:\n{e}")

    def select_buy(self):
        self.trade_type_var.set("Buy")
        self._update_trade_type_buttons()
        self.sync_all_sl_tp()

    def select_sell(self):
        self.trade_type_var.set("Sell")
        self._update_trade_type_buttons()
        self.sync_all_sl_tp()

    def _update_trade_type_buttons(self):
        buy_color_active = "#2ecc40"
        sell_color_active = "#ff4136"
        text_color_active = "#fff"
        default_bg = "#e0e0e0"
        default_fg = "#333"

        if self.trade_type_var.get() == "Buy":
            self.buy_btn.configure(
                bg=buy_color_active, fg=text_color_active,
                activebackground=buy_color_active, activeforeground=text_color_active,
                relief="solid", bd=1, highlightthickness=0,
                font=("Segoe UI", 11, "bold"), cursor="hand2"
            )
            self.sell_btn.configure(
                bg=default_bg, fg=default_fg,
                activebackground=default_bg, activeforeground=default_fg,
                relief="solid", bd=1, highlightthickness=0,
                font=("Segoe UI", 11), cursor="hand2"
            )
        else:
            self.sell_btn.configure(
                bg=sell_color_active, fg=text_color_active,
                activebackground=sell_color_active, activeforeground=text_color_active,
                relief="solid", bd=1, highlightthickness=0,
                font=("Segoe UI", 11, "bold"), cursor="hand2"
            )
            self.buy_btn.configure(
                bg=default_bg, fg=default_fg,
                activebackground=default_bg, activeforeground=default_fg,
                relief="solid", bd=1, highlightthickness=0,
                font=("Segoe UI", 11), cursor="hand2"
            )

    def update_balance_popup(self):
        popup = tk.Toplevel(self)
        popup.title("Update Account Balance")
        popup.geometry("280x120")
        popup.grab_set()
        popup.resizable(False, False)

        tk.Label(popup, text="Enter new account balance:").pack(pady=(14, 4))
        new_bal_var = tk.DoubleVar(value=self.account_balance_var.get())
        entry = ttk.Entry(popup, textvariable=new_bal_var, width=18)
        entry.pack(pady=4)
        entry.focus_set()
        self._on_focus_select_all(entry, new_bal_var)

        def update_and_close():
            try:
                self.account_balance_var.set(float(new_bal_var.get()))
                popup.destroy()
            except (ValueError, tk.TclError):
                entry.config(foreground="red")
                messagebox.showerror("Invalid Input", "Please enter a valid number for the balance.")
        ttk.Button(popup, text="Update", command=update_and_close).pack(pady=(10, 5))

    def update_sl_from_pips(self):
        if self._sl_tp_is_updating: return
        self._sl_tp_is_updating = True
        try:
            entry = self.entry_price_var.get()
            pips = self.sl_pips_var.get()

            if self.trade_type_var.get() == "Buy":
                price = entry - (pips * PIP_VALUE_XAUUSD)
            else:
                price = entry + (pips * PIP_VALUE_XAUUSD)

            if abs(self.sl_price_var.get() - price) > 1e-6:
                self.sl_price_var.set(round(price, 2))
        except (ValueError, tk.TclError):
            pass
        finally:
            self.update_sl_loss()
            self._sl_tp_is_updating = False

    def update_sl_from_price(self):
        if self._sl_tp_is_updating: return
        self._sl_tp_is_updating = True
        try:
            entry = self.entry_price_var.get()
            price = self.sl_price_var.get()

            if self.trade_type_var.get() == "Buy":
                pips = round((entry - price) / PIP_VALUE_XAUUSD, 1)
            else:
                pips = round((price - entry) / PIP_VALUE_XAUUSD, 1)

            if abs(self.sl_pips_var.get() - pips) > 1e-6:
                self.sl_pips_var.set(pips)
        except (ValueError, tk.TclError):
            pass
        finally:
            self.update_sl_loss()
            self._sl_tp_is_updating = False

    def update_tp_from_pips(self):
        if self._sl_tp_is_updating: return
        self._sl_tp_is_updating = True
        try:
            entry = self.entry_price_var.get()
            pips = self.tp_pips_var.get()

            if self.trade_type_var.get() == "Buy":
                price = entry + (pips * PIP_VALUE_XAUUSD)
            else:
                price = entry - (pips * PIP_VALUE_XAUUSD)

            if abs(self.tp_price_var.get() - price) > 1e-6:
                self.tp_price_var.set(round(price, 2))
        except (ValueError, tk.TclError):
            pass
        finally:
            self.update_tp_profit()
            self._sl_tp_is_updating = False

    def update_tp_from_price(self):
        if self._sl_tp_is_updating: return
        self._sl_tp_is_updating = True
        try:
            entry = self.entry_price_var.get()
            price = self.tp_price_var.get()

            if self.trade_type_var.get() == "Buy":
                pips = round((price - entry) / PIP_VALUE_XAUUSD, 1)
            else:
                pips = round((entry - price) / PIP_VALUE_XAUUSD, 1)

            if abs(self.tp_pips_var.get() - pips) > 1e-6:
                self.tp_pips_var.set(pips)
        except (ValueError, tk.TclError):
            pass
        finally:
            self.update_tp_profit()
            self._sl_tp_is_updating = False

    def update_sl_loss(self):
        try:
            lot = float(self.lot_size_var.get())
            pips = abs(float(self.sl_pips_var.get()))
            loss = -lot * pips * USD_PER_PIP_PER_LOT
            self.sl_loss_var.set("${:,.0f}".format(loss))
            balance = float(self.account_balance_var.get())
            loss_pct = (abs(loss) / balance) * 100 if balance > 0 else 0
            self.sl_loss_pct_var.set("({:.2f}%)".format(loss_pct))
        except (ValueError, tk.TclError):
            self.sl_loss_var.set("")
            self.sl_loss_pct_var.set("")
        except Exception:
            pass

    def update_tp_profit(self):
        try:
            lot = float(self.lot_size_var.get())
            pips = abs(float(self.tp_pips_var.get()))
            profit = lot * pips * USD_PER_PIP_PER_LOT
            self.tp_profit_var.set("${:,.0f}".format(profit))
            balance = float(self.account_balance_var.get())
            profit_pct = (profit / balance) * 100 if balance > 0 else 0
            self.tp_profit_pct_var.set("({:.2f}%)".format(profit_pct))
        except (ValueError, tk.TclError):
            self.tp_profit_var.set("")
            self.tp_profit_pct_var.set("")
        except Exception:
            pass

    def sync_all_sl_tp(self):
        self.update_sl_from_pips()
        self.update_tp_from_pips()

    def update_market_session(self, *args):
        try:
            self.timezone_display_var.set(self.timezone_var.get())
            tz = pytz.timezone(self.timezone_var.get())
        except Exception:
            tz = pytz.UTC

        try:
            date_str = self.trade_date_var.get()
            time_str = self.trade_time_var.get()
            if len(time_str.split(":")) == 1:
                time_str += ":00"

            local_dt = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
            local_dt = tz.localize(local_dt)
            utc_dt = local_dt.astimezone(pytz.UTC)
            utc_time = utc_dt.time()
        except ValueError:
            self.market_session_var.set("Invalid Date/Time")
            return
        except Exception:
            self.market_session_var.set("")
            return

        active_sessions = []
        overlaps = []
        for i, (name, start, end) in enumerate(MARKET_SESSIONS_UTC):
            if start < end:
                is_open = start <= utc_time < end
            else:
                is_open = utc_time >= start or utc_time < end
            if is_open:
                active_sessions.append(name)

        if len(active_sessions) > 1:
            for j in range(len(active_sessions)):
                for k in range(j + 1, len(active_sessions)):
                    overlaps.append(f"{active_sessions[j]}+{active_sessions[k]}")

        session_str = ""
        if overlaps:
            session_str = " / ".join(overlaps)
        elif active_sessions:
            session_str = " / ".join(active_sessions)
        else:
            session_str = "Closed"

        self.market_session_var.set(session_str)

if __name__ == "__main__":
    app = TradingJournalApp()
    app.mainloop()
