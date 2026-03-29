"""
RoadtripDB - GUI for logging road trip data into a Polars DataFrame.

Fields: Travel Time (min), Stop Time (min), Distance (mi), Place Name
Data is stored in a Polars DataFrame and displayed in a table view.
Arrival and Departure times are computed cumulatively from the trip start time.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
import calendar
import json
import polars as pl


class RoadtripApp:
    """Main application window for the Roadtrip data entry GUI."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("RoadtripDB")
        self.root.geometry("1060x660")
        self.root.minsize(900, 520)
        self.root.configure(bg="#1e1e2e")

        # Initialize an empty Polars DataFrame
        self.df = pl.DataFrame(
            schema={
                "Place Name": pl.Utf8,
                "Travel Time": pl.Utf8,
                "Stop Time": pl.Utf8,
                "Distance (mi)": pl.Float64,
                "Arrival Time": pl.Utf8,
                "Departure Time": pl.Utf8,
            }
        )

        # Map treeview column IDs to DataFrame column names
        self._col_map = {
            "place": "Place Name",
            "travel": "Travel Time",
            "stop": "Stop Time",
            "distance": "Distance (mi)",
        }
        # Columns that the user may NOT edit (computed)
        self._readonly_cols = {"stopnum", "arrival", "departure"}
        # Track the active edit widget
        self._edit_widget: tk.Entry | None = None
        # Track form edit index for the main edit button feature
        self._editing_idx: int | None = None
        # Dirty flag — set when data changes, cleared on save/export
        self._dirty = False

        self._apply_theme()
        self._build_ui()

        # Intercept window close to check for unsaved changes
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Theme ───────────────────────────────────────────────────────────
    def _apply_theme(self) -> None:
        """Configure ttk styles for a modern dark theme."""
        style = ttk.Style(self.root)
        style.theme_use("clam")

        # Color palette
        bg = "#1e1e2e"
        surface = "#2a2a3c"
        accent = "#7c3aed"
        accent_hover = "#6d28d9"
        text = "#e2e2f0"
        muted = "#9090a8"
        border = "#3a3a50"

        style.configure(".", background=bg, foreground=text, font=("Helvetica", 15))

        # Title label
        style.configure("Title.TLabel", background=bg, foreground=text,
                         font=("Helvetica", 23, "bold"))

        # Subtitle / section labels
        style.configure("Section.TLabel", background=bg, foreground=muted,
                         font=("Helvetica", 14))

        # Entry fields
        style.configure("Custom.TEntry", fieldbackground=surface,
                         foreground=text, bordercolor=border,
                         insertcolor=text, padding=8)
        style.map("Custom.TEntry",
                   bordercolor=[("focus", accent)],
                   lightcolor=[("focus", accent)])

        # Buttons
        style.configure("Accent.TButton", background=accent,
                         foreground="#ffffff", font=("Helvetica", 15, "bold"),
                         padding=(16, 10), borderwidth=0)
        style.map("Accent.TButton",
                   background=[("active", accent_hover), ("pressed", accent_hover)])

        style.configure("Clear.TButton", background="#ef4444",
                         foreground="#ffffff", font=("Helvetica", 15, "bold"),
                         padding=(16, 10), borderwidth=0)
        style.map("Clear.TButton",
                   background=[("active", "#dc2626"), ("pressed", "#dc2626")])

        # Treeview (table)
        style.configure("Treeview", background=surface, foreground=text,
                         fieldbackground=surface, borderwidth=0, rowheight=36,
                         font=("Helvetica", 14))
        style.configure("Treeview.Heading", background=border,
                         foreground=text, font=("Helvetica", 14, "bold"),
                         padding=6)
        style.map("Treeview",
                   background=[("selected", accent)],
                   foreground=[("selected", "#ffffff")])

        # Frame
        style.configure("Card.TFrame", background=surface, relief="flat")

        # Label frame
        style.configure("Card.TLabelframe", background=surface,
                         foreground=text, bordercolor=border)
        style.configure("Card.TLabelframe.Label", background=surface,
                         foreground=muted, font=("Helvetica", 14, "bold"))

        # Separator
        style.configure("TSeparator", background=border)

        # Status bar
        style.configure("Status.TLabel", background=bg, foreground=muted,
                         font=("Helvetica", 13))

    # ── UI ──────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        """Construct the full user interface."""
        # Title
        title = ttk.Label(self.root, text="🗺  RoadtripDB", style="Title.TLabel")
        title.pack(pady=(18, 2))

        subtitle = ttk.Label(self.root, text="Log your road trip stops",
                              style="Section.TLabel")
        subtitle.pack(pady=(0, 12))

        ttk.Separator(self.root).pack(fill="x", padx=24)

        # ── Start Time Section ──────────────────────────────────────────
        start_frame = ttk.Frame(self.root)
        start_frame.pack(fill="x", padx=24, pady=(10, 0))

        now = datetime.now().replace(second=0, microsecond=0)

        ttk.Label(start_frame, text="Trip Start",
                   style="Section.TLabel").pack(side="left", padx=(0, 12))

        # Shared spinbox style kwargs
        spin_kw = dict(
            font=("Helvetica", 15), justify="center",
            bg="#2a2a3c", fg="#e2e2f0", insertbackground="#e2e2f0",
            highlightthickness=1, highlightbackground="#3a3a50",
            highlightcolor="#7c3aed", buttonbackground="#3a3a50",
            relief="flat", wrap=True,
        )

        # Month
        ttk.Label(start_frame, text="Month",
                   style="Section.TLabel").pack(side="left", padx=(0, 3))
        self.start_month = tk.Spinbox(
            start_frame, from_=1, to=12, width=3,
            format="%02.0f", **spin_kw,
        )
        self.start_month.delete(0, "end")
        self.start_month.insert(0, f"{now.month:02d}")
        self.start_month.pack(side="left", padx=(0, 4))

        ttk.Label(start_frame, text="/",
                   style="Section.TLabel").pack(side="left")

        # Day
        ttk.Label(start_frame, text="Day",
                   style="Section.TLabel").pack(side="left", padx=(4, 3))
        self.start_day = tk.Spinbox(
            start_frame, from_=1, to=31, width=3,
            format="%02.0f", **spin_kw,
        )
        self.start_day.delete(0, "end")
        self.start_day.insert(0, f"{now.day:02d}")
        self.start_day.pack(side="left", padx=(0, 4))

        ttk.Label(start_frame, text="/",
                   style="Section.TLabel").pack(side="left")

        # Year
        ttk.Label(start_frame, text="Year",
                   style="Section.TLabel").pack(side="left", padx=(4, 3))
        self.start_year = tk.Spinbox(
            start_frame, from_=2020, to=2040, width=5, **spin_kw,
        )
        self.start_year.delete(0, "end")
        self.start_year.insert(0, str(now.year))
        self.start_year.pack(side="left", padx=(0, 16))

        # Hour
        ttk.Label(start_frame, text="Hour",
                   style="Section.TLabel").pack(side="left", padx=(0, 3))
        self.start_hour = tk.Spinbox(
            start_frame, from_=0, to=23, width=3,
            format="%02.0f", **spin_kw,
        )
        self.start_hour.delete(0, "end")
        self.start_hour.insert(0, f"{now.hour:02d}")
        self.start_hour.pack(side="left", padx=(0, 4))

        ttk.Label(start_frame, text=":",
                   style="Section.TLabel").pack(side="left")

        # Minute
        ttk.Label(start_frame, text="Min",
                   style="Section.TLabel").pack(side="left", padx=(4, 3))
        self.start_minute = tk.Spinbox(
            start_frame, from_=0, to=59, width=3,
            format="%02.0f", **spin_kw,
        )
        self.start_minute.delete(0, "end")
        self.start_minute.insert(0, f"{now.minute:02d}")
        self.start_minute.pack(side="left")

        # ── Input Section ───────────────────────────────────────────────
        input_frame = ttk.LabelFrame(self.root, text="  New Entry  ",
                                      style="Card.TLabelframe", padding=16)
        input_frame.pack(fill="x", padx=24, pady=(12, 6))

        # Grid of labels + entries
        fields = [
            ("Place Name", "e.g. Grand Canyon"),
            ("Distance (mi)", "e.g. 85.5"),
            ("Travel Time", "e.g. 2:30"),
            ("Stop Time", "e.g. 0:45"),
        ]

        self.entries: dict[str, ttk.Entry] = {}

        for col, (label_text, placeholder) in enumerate(fields):
            lbl = ttk.Label(input_frame, text=label_text,
                             style="Section.TLabel")
            lbl.grid(row=0, column=col, padx=8, pady=(0, 4), sticky="w")

            entry = ttk.Entry(input_frame, style="Custom.TEntry", width=18)
            entry.grid(row=1, column=col, padx=8, sticky="ew")
            entry.insert(0, placeholder)
            entry.configure(foreground="#6a6a80")
            entry.bind("<FocusIn>", lambda e, ent=entry, ph=placeholder:
                        self._on_focus_in(ent, ph))
            entry.bind("<FocusOut>", lambda e, ent=entry, ph=placeholder:
                        self._on_focus_out(ent, ph))
            self.entries[label_text] = entry

        input_frame.columnconfigure((0, 1, 2, 3), weight=1)

        # Bind Enter key to add entry
        self.root.bind("<Return>", self._handle_return)

        # ── Buttons ─────────────────────────────────────────────────────
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=24, pady=8)

        self.action_frame = ttk.Frame(btn_frame)
        self.action_frame.pack(side="left", padx=(0, 8))

        self.add_btn = ttk.Button(self.action_frame, text="＋  Add Entry",
                              style="Accent.TButton", command=self._add_entry)
        self.add_btn.pack(side="left")

        self.cancel_edit_btn = ttk.Button(self.action_frame, text="✕  Cancel",
                                 style="Clear.TButton", command=self._cancel_edit_entry)

        self.insert_btn = ttk.Button(btn_frame, text="⤵  Insert",
                                 style="Accent.TButton", command=self._insert_entry,
                                 state="disabled")
        self.insert_btn.pack(side="left", padx=(0, 8))

        self.edit_btn = ttk.Button(btn_frame, text="✏️  Edit",
                                style="Accent.TButton", command=self._start_edit_entry,
                                state="disabled")
        self.edit_btn.pack(side="left", padx=(0, 8))

        self.delete_btn = ttk.Button(btn_frame, text="🗑  Delete",
                                 style="Clear.TButton", command=self._delete_entry,
                                 state="disabled")
        self.delete_btn.pack(side="left", padx=(0, 8))

        self.clear_btn = ttk.Button(btn_frame, text="✕  Clear All",
                                style="Clear.TButton", command=self._clear_all,
                                state="disabled")
        self.clear_btn.pack(side="left", padx=(0, 8))

        self.import_btn = ttk.Button(btn_frame, text="📂  Load Trip",
                                  style="Accent.TButton", command=self._import_trip)
        self.import_btn.pack(side="right", padx=(0, 8))

        self.export_btn = ttk.Button(btn_frame, text="💾  Save Trip",
                                 style="Accent.TButton", command=self._export_trip,
                                 state="disabled")
        self.export_btn.pack(side="right")

        # ── Table ───────────────────────────────────────────────────────
        table_frame = ttk.Frame(self.root, style="Card.TFrame")
        table_frame.pack(fill="both", expand=True, padx=24, pady=(4, 8))

        columns = ("stopnum", "place", "distance", "travel", "arrival", "stop", "departure")
        self.tree = ttk.Treeview(table_frame, columns=columns,
                                  show="headings", selectmode="browse")

        self.tree.heading("stopnum", text="#")
        self.tree.heading("place", text="Place Name")
        self.tree.heading("distance", text="Distance (mi)")
        self.tree.heading("travel", text="Travel Time")
        self.tree.heading("arrival", text="Arrival Time")
        self.tree.heading("stop", text="Stop Time")
        self.tree.heading("departure", text="Departure Time")

        self.tree.column("stopnum", width=45, anchor="center", stretch=False)
        self.tree.column("place", width=180, anchor="w")
        self.tree.column("distance", width=120, anchor="center")
        self.tree.column("travel", width=120, anchor="center")
        self.tree.column("arrival", width=150, anchor="center")
        self.tree.column("stop", width=120, anchor="center")
        self.tree.column("departure", width=150, anchor="center")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical",
                                   command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Right-click to delete row
        self.tree.bind("<Button-2>", self._delete_selected)  # macOS right-click
        self.tree.bind("<Button-3>", self._delete_selected)  # Windows/Linux

        # Double-click to edit a cell
        self.tree.bind("<Double-1>", self._on_double_click)

        # ── Status Bar ──────────────────────────────────────────────────
        self.status_var = tk.StringVar(
            value="0 entries  |  Double-click to edit  |  Right-click to delete"
        )
        status_bar = ttk.Label(self.root, textvariable=self.status_var,
                                style="Status.TLabel")
        status_bar.pack(fill="x", padx=24, pady=(0, 10))

        # Default focus to the Load Trip button
        self.import_btn.focus_set()

    # ── Placeholder helpers ─────────────────────────────────────────────
    @staticmethod
    def _on_focus_in(entry: ttk.Entry, placeholder: str) -> None:
        if entry.get() == placeholder:
            entry.delete(0, "end")
            entry.configure(foreground="#e2e2f0")

    @staticmethod
    def _on_focus_out(entry: ttk.Entry, placeholder: str) -> None:
        if not entry.get():
            entry.insert(0, placeholder)
            entry.configure(foreground="#6a6a80")

    # ── Actions ─────────────────────────────────────────────────────────
    def _validate_inputs(self) -> tuple[dict[str, str], float] | None:
        """Validate the input fields and return (raw_values, distance) or None."""
        placeholders = {
            "Place Name": "e.g. Grand Canyon",
            "Travel Time": "e.g. 2:30",
            "Stop Time": "e.g. 0:45",
            "Distance (mi)": "e.g. 85.5",
        }

        raw: dict[str, str] = {}
        for key, entry in self.entries.items():
            val = entry.get().strip()
            if val == placeholders[key] or val == "":
                messagebox.showwarning("Missing field",
                                        f'Please fill in "{key}".')
                entry.focus_set()
                return None
            raw[key] = val

        # Validate H:MM time fields
        if self._parse_hm(raw["Travel Time"]) is None:
            messagebox.showerror("Invalid input",
                                  'Travel Time must be in H:MM format (e.g. 2:30).')
            return None
        if self._parse_hm(raw["Stop Time"]) is None:
            messagebox.showerror("Invalid input",
                                  'Stop Time must be in H:MM format (e.g. 0:45).')
            return None

        try:
            dist = float(raw["Distance (mi)"])
        except ValueError:
            messagebox.showerror("Invalid input", "Distance must be a number.")
            return None

        return raw, dist

    def _make_new_row(self, raw: dict[str, str], dist: float) -> pl.DataFrame:
        """Build a single-row DataFrame from validated inputs."""
        return pl.DataFrame({
            "Place Name": [raw["Place Name"]],
            "Travel Time": [raw["Travel Time"]],
            "Stop Time": [raw["Stop Time"]],
            "Distance (mi)": [dist],
            "Arrival Time": [""],
            "Departure Time": [""],
        })

    def _clear_inputs(self) -> None:
        """Reset all input fields to their placeholders."""
        placeholders = {
            "Place Name": "e.g. Grand Canyon",
            "Travel Time": "e.g. 2:30",
            "Stop Time": "e.g. 0:45",
            "Distance (mi)": "e.g. 85.5",
        }
        for key, entry in self.entries.items():
            entry.delete(0, "end")
            entry.insert(0, placeholders[key])
            entry.configure(foreground="#6a6a80")

    def _handle_return(self, event: tk.Event) -> None:
        if self._editing_idx is not None:
            self._save_edit_entry()
        else:
            self._add_entry()

    def _start_edit_entry(self) -> None:
        """Start editing the selected row using the main input form."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("No row selected", "Select a row in the table to edit.")
            return

        item = selected[0]
        self._editing_idx = self.tree.index(item)

        # Clear inputs to get placeholders, then fill with row data
        self._clear_inputs()
        row = self.df.to_dicts()[self._editing_idx]

        for key, entry in self.entries.items():
            entry.delete(0, "end")
            val = str(row[key])
            entry.insert(0, val)
            if val:
                entry.configure(foreground="#e2e2f0")

        self.add_btn.configure(text="✓  Save Edit", command=self._save_edit_entry)
        self.cancel_edit_btn.pack(side="left", padx=(8, 0))
        self.entries["Place Name"].focus_set()
        self._update_status()

    def _save_edit_entry(self) -> None:
        """Save form changes back to the DataFrame."""
        if self._editing_idx is None:
            return

        result = self._validate_inputs()
        if result is None:
            return
        
        raw, dist = result
        new_row = self._make_new_row(raw, dist)

        top = self.df.slice(0, self._editing_idx)
        bottom = self.df.slice(self._editing_idx + 1)
        self.df = pl.concat([top, new_row, bottom])

        self._recalculate_times()
        self._refresh_treeview()
        self._clear_inputs()

        self._dirty = True
        self._cancel_edit_entry(clear=False)
        self._update_status()

        print("\n── Saved Edit to DataFrame ───────────────────────")
        print(self.df)
        print("──────────────────────────────────────────────────\n")

    def _cancel_edit_entry(self, clear: bool = True) -> None:
        """Cancel the current edit or return the UI state form after save."""
        self._editing_idx = None
        if clear:
            self._clear_inputs()
        
        self.add_btn.configure(text="＋  Add Entry", command=self._add_entry)
        self.cancel_edit_btn.pack_forget()
        self._update_status()

    def _add_entry(self) -> None:
        """Validate inputs and append a new row to the DataFrame + table."""
        result = self._validate_inputs()
        if result is None:
            return
        raw, dist = result

        self.df = pl.concat([self.df, self._make_new_row(raw, dist)])

        self._recalculate_times()
        self._refresh_treeview()
        self._clear_inputs()

        self._dirty = True
        self._update_status()
        self.entries["Place Name"].focus_set()

        print("\n── Current DataFrame ─────────────────────────────")
        print(self.df)
        print("──────────────────────────────────────────────────\n")

    def _insert_entry(self) -> None:
        """Validate inputs and insert a new row above the selected row."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("No row selected",
                                 "Select a row in the table first.\n"
                                 "The new entry will be inserted above it.")
            return

        result = self._validate_inputs()
        if result is None:
            return
        raw, dist = result

        insert_idx = self.tree.index(selected[0])
        new_row = self._make_new_row(raw, dist)

        # Split the DataFrame and sandwich the new row in
        top = self.df.slice(0, insert_idx)
        bottom = self.df.slice(insert_idx)
        self.df = pl.concat([top, new_row, bottom])

        self._recalculate_times()
        self._refresh_treeview()
        self._clear_inputs()

        self._dirty = True
        self._update_status()
        self.entries["Place Name"].focus_set()

        print("\n── Inserted into DataFrame ───────────────────────")
        print(self.df)
        print("──────────────────────────────────────────────────\n")

    def _delete_entry(self) -> None:
        """Delete the selected row via the toolbar button."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("No row selected",
                                 "Select a row in the table to delete.")
            return

        item = selected[0]
        idx = self.tree.index(item)
        confirm = messagebox.askyesno("Delete entry",
                                       f"Delete row {idx + 1}?")
        if confirm:
            self.tree.delete(item)
            mask = [True] * len(self.df)
            mask[idx] = False
            self.df = self.df.filter(pl.Series(mask))
            self._recalculate_times()
            self._refresh_treeview()
            self._dirty = True
            self._update_status()
            print("\n── Deleted from DataFrame ────────────────────────")
            print(self.df)
            print("──────────────────────────────────────────────────\n")

    def _delete_selected(self, event: tk.Event) -> None:
        """Delete the selected row from the table and DataFrame."""
        item = self.tree.identify_row(event.y)
        if not item:
            return

        idx = self.tree.index(item)
        confirm = messagebox.askyesno("Delete entry",
                                       f"Delete row {idx + 1}?")
        if confirm:
            self.tree.delete(item)
            # Rebuild DataFrame without the deleted row
            mask = [True] * len(self.df)
            mask[idx] = False
            self.df = self.df.filter(pl.Series(mask))
            # Recalculate times after deletion
            self._recalculate_times()
            self._refresh_treeview()
            self._dirty = True
            self._update_status()
            print("\n── Updated DataFrame ─────────────────────────────")
            print(self.df)
            print("──────────────────────────────────────────────────\n")

    def _clear_all(self) -> None:
        """Clear all entries from the DataFrame and table."""
        if len(self.df) == 0:
            return
        confirm = messagebox.askyesno("Clear all",
                                       "Remove all entries?")
        if confirm:
            self.df = self.df.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            self._dirty = False
            self._update_status()

    def _export_trip(self) -> None:
        """Save the trip (start time + stops) to a JSON file."""
        if len(self.df) == 0:
            messagebox.showinfo("Nothing to save",
                                 "Add some entries first.")
            return

        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("Trip files", "*.json"), ("All files", "*.*")],
            initialfile="roadtrip.json",
        )
        if not path:
            return

        # Build the trip start string from spinboxes
        try:
            start_dt = self._parse_start_time()
            start_str = start_dt.strftime("%Y-%m-%d %H:%M")
        except ValueError:
            return

        # Convert DataFrame rows to list of dicts
        stops = self.df.to_dicts()

        trip_data = {
            "trip_start": start_str,
            "stops": stops,
        }

        with open(path, "w") as f:
            json.dump(trip_data, f, indent=2)

        self._dirty = False
        self._update_status()
        messagebox.showinfo("Saved",
                             f"Trip saved ({len(self.df)} stops) to:\n{path}")

    def _import_trip(self) -> None:
        """Load a trip (start time + stops) from a JSON file."""
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            filetypes=[
                ("Trip files", "*.json"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        try:
            with open(path, "r") as f:
                trip_data = json.load(f)
        except Exception as exc:
            messagebox.showerror("Load failed", f"Could not read file:\n{exc}")
            return

        # Validate structure
        if "stops" not in trip_data or not isinstance(trip_data["stops"], list):
            messagebox.showerror("Invalid file",
                                  "File does not contain valid trip data.")
            return

        # Restore trip start time
        if "trip_start" in trip_data:
            try:
                dt = datetime.strptime(trip_data["trip_start"], "%Y-%m-%d %H:%M")
                self.start_month.delete(0, "end")
                self.start_month.insert(0, f"{dt.month:02d}")
                self.start_day.delete(0, "end")
                self.start_day.insert(0, f"{dt.day:02d}")
                self.start_year.delete(0, "end")
                self.start_year.insert(0, str(dt.year))
                self.start_hour.delete(0, "end")
                self.start_hour.insert(0, f"{dt.hour:02d}")
                self.start_minute.delete(0, "end")
                self.start_minute.insert(0, f"{dt.minute:02d}")
            except ValueError:
                pass  # keep current start time if parse fails

        # Build DataFrame from stops
        try:
            imported = pl.DataFrame(trip_data["stops"])
        except Exception as exc:
            messagebox.showerror("Load failed",
                                  f"Could not parse stop data:\n{exc}")
            return

        # Validate required columns
        required = {"Place Name", "Travel Time", "Stop Time", "Distance (mi)"}
        missing = required - set(imported.columns)
        if missing:
            messagebox.showerror(
                "Invalid data",
                f"Missing required fields:\n{', '.join(sorted(missing))}",
            )
            return

        # Ensure correct types
        try:
            imported = imported.with_columns([
                pl.col("Travel Time").cast(pl.Utf8),
                pl.col("Stop Time").cast(pl.Utf8),
                pl.col("Distance (mi)").cast(pl.Float64),
            ])
        except Exception:
            messagebox.showerror(
                "Invalid data",
                "Could not parse columns. Travel/Stop Time should be H:MM,\n"
                "Distance should be a number.",
            )
            return

        # Add Arrival / Departure columns if not already present
        if "Arrival Time" not in imported.columns:
            imported = imported.with_columns(pl.lit("").alias("Arrival Time"))
        if "Departure Time" not in imported.columns:
            imported = imported.with_columns(pl.lit("").alias("Departure Time"))

        # Keep only the columns we care about, in the right order
        imported = imported.select(
            "Place Name", "Travel Time", "Stop Time",
            "Distance (mi)", "Arrival Time", "Departure Time",
        )

        self.df = imported
        self._recalculate_times()
        self._refresh_treeview()
        self._dirty = False
        self._update_status()

        messagebox.showinfo("Loaded",
                             f"Trip loaded ({len(self.df)} stops) from:\n{path}")

        print("\n── Loaded Trip ───────────────────────────────────")
        print(f"Start: {trip_data.get('trip_start', 'N/A')}")
        print(self.df)
        print("──────────────────────────────────────────────────\n")

    # ── Time Computation ────────────────────────────────────────────────
    def _parse_start_time(self) -> datetime:
        """Read the trip start time from the spinbox widgets."""
        try:
            year = int(self.start_year.get())
            month = int(self.start_month.get())
            day = int(self.start_day.get())
            hour = int(self.start_hour.get())
            minute = int(self.start_minute.get())
            # Clamp day to max days in that month
            max_day = calendar.monthrange(year, month)[1]
            day = min(day, max_day)
            return datetime(year, month, day, hour, minute)
        except (ValueError, OverflowError):
            messagebox.showerror(
                "Invalid start time",
                "Please set a valid start date and time."
            )
            raise ValueError("Invalid start time")

    @staticmethod
    def _parse_hm(value: str) -> int | None:
        """Parse an 'H:MM' string and return total minutes, or None on failure."""
        value = value.strip()
        if ":" not in value:
            return None
        parts = value.split(":", 1)
        try:
            hours = int(parts[0])
            mins = int(parts[1])
        except ValueError:
            return None
        if mins < 0 or mins > 59 or hours < 0:
            return None
        return hours * 60 + mins

    @staticmethod
    def _format_hm(total_minutes: int) -> str:
        """Convert total minutes to 'H:MM' string."""
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h}:{m:02d}"

    def _recalculate_times(self) -> None:
        """Recompute Arrival and Departure for every row based on start time."""
        if len(self.df) == 0:
            return

        try:
            cursor = self._parse_start_time()
        except ValueError:
            return

        arrivals: list[str] = []
        departures: list[str] = []

        for row in self.df.iter_rows(named=True):
            travel_min = self._parse_hm(row["Travel Time"]) or 0
            stop_min = self._parse_hm(row["Stop Time"]) or 0

            arrival = cursor + timedelta(minutes=travel_min)
            departure = arrival + timedelta(minutes=stop_min)

            arrivals.append(arrival.strftime("%Y-%m-%d %H:%M"))
            departures.append(departure.strftime("%Y-%m-%d %H:%M"))

            cursor = departure  # next leg starts when we depart

        self.df = self.df.with_columns([
            pl.Series("Arrival Time", arrivals),
            pl.Series("Departure Time", departures),
        ])

    def _refresh_treeview(self) -> None:
        """Rebuild the treeview from the current DataFrame."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, row in enumerate(self.df.iter_rows(named=True), start=1):
            self.tree.insert("", "end", values=(
                idx,
                row["Place Name"],
                row["Distance (mi)"],
                row["Travel Time"],
                row["Arrival Time"],
                row["Stop Time"],
                row["Departure Time"],
            ))

    # ── Inline Cell Editing ──────────────────────────────────────────────
    def _on_double_click(self, event: tk.Event) -> None:
        """Show an Entry widget over the double-clicked cell for editing."""
        # Dismiss any existing edit widget first
        self._cancel_edit()

        # Identify which row and column were clicked
        item = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)  # e.g. "#1", "#2", ...
        if not item or not col_id:
            return

        # Convert "#N" to the column key
        col_index = int(col_id.replace("#", "")) - 1
        col_key = self.tree["columns"][col_index]  # e.g. "place", "travel"

        # Block editing on computed columns
        if col_key in self._readonly_cols:
            return

        # Get the cell bounding box (x, y, width, height) relative to the treeview
        bbox = self.tree.bbox(item, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        # Current cell value
        current_values = self.tree.item(item, "values")
        current_text = str(current_values[col_index])

        # Create an overlay Entry widget
        entry = tk.Entry(
            self.tree,
            font=("Helvetica", 14),
            bg="#2a2a3c",
            fg="#e2e2f0",
            insertbackground="#e2e2f0",
            selectbackground="#7c3aed",
            selectforeground="#ffffff",
            highlightthickness=2,
            highlightcolor="#7c3aed",
            highlightbackground="#7c3aed",
            relief="flat",
        )
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, current_text)
        entry.select_range(0, "end")
        entry.focus_set()

        # Store context for commit
        self._edit_widget = entry
        self._edit_item = item
        self._edit_col_key = col_key
        self._edit_col_index = col_index

        # Bind confirm / cancel
        entry.bind("<Return>", lambda e: self._commit_edit())
        entry.bind("<Escape>", lambda e: self._cancel_edit())
        entry.bind("<FocusOut>", lambda e: self._commit_edit())

    def _commit_edit(self) -> None:
        """Write the edited value back to the DataFrame and recalculate."""
        if self._edit_widget is None:
            return

        new_value = self._edit_widget.get().strip()
        item = self._edit_item
        col_key = self._edit_col_key
        col_index = self._edit_col_index
        row_idx = self.tree.index(item)

        # Destroy the widget immediately to prevent re-entrant FocusOut
        widget = self._edit_widget
        self._edit_widget = None
        widget.destroy()

        # Determine the DataFrame column name
        df_col = self._col_map[col_key]

        # Validate based on column type
        if col_key in ("travel", "stop"):
            # H:MM time fields
            if self._parse_hm(new_value) is None:
                messagebox.showerror("Invalid input",
                                      f'"{df_col}" must be in H:MM format (e.g. 1:30).')
                return
            col_values = self.df[df_col].to_list()
            col_values[row_idx] = new_value
            self.df = self.df.with_columns(pl.Series(df_col, col_values))
        elif col_key == "distance":
            try:
                numeric_val = float(new_value)
            except ValueError:
                messagebox.showerror("Invalid input",
                                      f'"{df_col}" must be a number.')
                return
            col_values = self.df[df_col].to_list()
            col_values[row_idx] = numeric_val
            self.df = self.df.with_columns(pl.Series(df_col, col_values))
        else:
            # Text column (Place Name)
            if not new_value:
                messagebox.showwarning("Empty value", "Place Name cannot be empty.")
                return
            col_values = self.df[df_col].to_list()
            col_values[row_idx] = new_value
            self.df = self.df.with_columns(pl.Series(df_col, col_values))

        # Recalculate arrival/departure and refresh
        self._recalculate_times()
        self._refresh_treeview()
        self._dirty = True
        self._update_status()

        print("\n── Edited DataFrame ──────────────────────────────")
        print(self.df)
        print("──────────────────────────────────────────────────\n")

    def _cancel_edit(self) -> None:
        """Dismiss the edit widget without saving."""
        if self._edit_widget is not None:
            widget = self._edit_widget
            self._edit_widget = None
            widget.destroy()

    def _update_status(self) -> None:
        """Refresh the status bar text and toggle button states."""
        n = len(self.df)
        word = "entry" if n == 1 else "entries"
        dirty_marker = "  •  Unsaved changes" if self._dirty else ""
        editing_marker = "  |  Editing row" if getattr(self, "_editing_idx", None) is not None else ""
        self.status_var.set(
            f"{n} {word}  |  Double-click to edit  |  Right-click to delete{dirty_marker}{editing_marker}"
        )

        # Enable / disable buttons that require grid data
        if getattr(self, "_editing_idx", None) is not None:
            state = "disabled"
        else:
            state = "!disabled" if n > 0 else "disabled"
            
        for btn in (self.insert_btn, self.edit_btn, self.delete_btn, self.clear_btn, self.export_btn):
            btn.state([state])

    def _on_close(self) -> None:
        """Handle window close — prompt to save if there are unsaved changes."""
        if self._dirty and len(self.df) > 0:
            answer = messagebox.askyesnocancel(
                "Unsaved changes",
                "You have unsaved changes.\n\nDo you want to export before closing?",
            )
            if answer is None:
                # Cancel — do not close
                return
            if answer:
                # Yes — export first, then close (unless user cancels the dialog)
                self._export_trip()
                if self._dirty:
                    # User cancelled the save dialog — stay open
                    return
        self.root.destroy()


# ── Main ────────────────────────────────────────────────────────────────
def main() -> None:
    root = tk.Tk()
    RoadtripApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
