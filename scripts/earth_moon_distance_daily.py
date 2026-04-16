#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import os
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from skyfield.api import load, utc
except ImportError:
    print("Missing dependency: skyfield", file=sys.stderr)
    print("Install with: pip install skyfield", file=sys.stderr)
    raise SystemExit(1)


KM_TO_MILES = 0.621371
CSV_HEADER = ["date_utc", "distance_km", "distance_miles"]


@dataclass(frozen=True)
class DistanceRow:
    day: date
    distance_km: float
    x_km: float = 0.0
    y_km: float = 0.0

    @property
    def distance_miles(self) -> float:
        return self.distance_km * KM_TO_MILES


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Measure Earth-Moon distance for each day in a date range.",
        epilog=(
            "Examples:\n"
            "  python earth_moon_distance_daily.py\n"
            "  python earth_moon_distance_daily.py --start-date 2026-04-16 --days 60 --output moon.csv\n"
            "  python earth_moon_distance_daily.py --time-utc 12:00 --plot\n"
            "  python earth_moon_distance_daily.py --append-daily"
        ),
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--start-date",
        default=date.today().isoformat(),
        help="Start date in ISO format YYYY-MM-DD (default: today)",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=30,
        help="Number of days to calculate (default: 30)",
    )
    parser.add_argument(
        "--time-utc",
        default="00:00",
        help="UTC time used for each day, format HH:MM (default: 00:00)",
    )
    parser.add_argument(
        "--output",
        default="earth_moon_distance.csv",
        help="Output CSV file path (default: earth_moon_distance.csv)",
    )
    parser.add_argument(
        "--plot",
        action="store_true",
        help="Create a PNG distance plot",
    )
    parser.add_argument(
        "--plot-output",
        default="",
        help="Plot PNG path (default: same path as CSV with .png extension)",
    )
    parser.add_argument(
        "--append-daily",
        action="store_true",
        help="Append one row for --start-date (or today) to the CSV, if not present",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch a simple desktop GUI",
    )
    return parser.parse_args()


def parse_iso_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Invalid --start-date: {value}") from exc


def parse_utc_time(value: str) -> time:
    try:
        return time.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Invalid --time-utc: {value}") from exc


def measure_distances(start_day: date, days: int, utc_time: time) -> list[DistanceRow]:
    if days <= 0:
        raise ValueError("--days must be greater than zero")

    ts = load.timescale()
    eph = load("de421.bsp")
    earth = eph["earth"]
    moon = eph["moon"]

    rows: list[DistanceRow] = []
    for offset in range(days):
        day = start_day + timedelta(days=offset)
        dt = datetime.combine(day, utc_time).replace(tzinfo=utc)
        t = ts.from_datetime(dt)
        observation = earth.at(t).observe(moon)
        distance_km = float(observation.distance().km)
        moon_position_km = observation.position.km
        rows.append(
            DistanceRow(
                day=day,
                distance_km=distance_km,
                x_km=float(moon_position_km[0]),
                y_km=float(moon_position_km[1]),
            )
        )

    return rows


def write_csv(rows: list[DistanceRow], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as fp:
        writer = csv.writer(fp)
        writer.writerow(CSV_HEADER)
        for row in rows:
            writer.writerow(
                [
                    row.day.isoformat(),
                    f"{row.distance_km:.3f}",
                    f"{row.distance_miles:.3f}",
                ]
            )


def read_existing_dates(output_path: Path) -> set[str]:
    if not output_path.exists():
        return set()

    dates: set[str] = set()
    with output_path.open("r", newline="", encoding="utf-8") as fp:
        reader = csv.DictReader(fp)
        for row in reader:
            day_value = (row.get("date_utc") or "").strip()
            if day_value:
                dates.add(day_value)
    return dates


def append_csv(rows: Iterable[DistanceRow], output_path: Path) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    existing_dates = read_existing_dates(output_path)
    write_header = not output_path.exists()
    appended = 0

    with output_path.open("a", newline="", encoding="utf-8") as fp:
        writer = csv.writer(fp)
        if write_header:
            writer.writerow(CSV_HEADER)

        for row in rows:
            day_key = row.day.isoformat()
            if day_key in existing_dates:
                continue

            writer.writerow(
                [
                    day_key,
                    f"{row.distance_km:.3f}",
                    f"{row.distance_miles:.3f}",
                ]
            )
            existing_dates.add(day_key)
            appended += 1

    return appended


def create_plot(rows: list[DistanceRow], plot_output: Path, utc_time: time, max_ticks: int = 10) -> None:
    try:
        import matplotlib.pyplot as plt
        import matplotlib.dates as mdates
    except ImportError as exc:
        raise RuntimeError("Missing dependency: matplotlib (pip install matplotlib)") from exc

    plot_output.parent.mkdir(parents=True, exist_ok=True)
    dates = [datetime.combine(row.day, time.min) for row in rows]
    distances = [row.distance_km for row in rows]
    min_row = min(rows, key=lambda item: item.distance_km)
    max_row = max(rows, key=lambda item: item.distance_km)

    fig, (ax_distance, ax_linear) = plt.subplots(1, 2, figsize=(14, 5))
    ax_distance.plot(dates, distances, marker="o", linewidth=1.4, markersize=3)
    ax_distance.set_title("Earth-Moon Distance by Day (km)")
    ax_distance.set_xlabel("Date (UTC)")
    ax_distance.set_ylabel("Distance (km)")
    ax_distance.grid(alpha=0.35)

    bounded_max_ticks = max(4, min(20, int(max_ticks)))
    locator = mdates.AutoDateLocator(minticks=max(3, bounded_max_ticks // 2), maxticks=bounded_max_ticks)
    formatter = mdates.ConciseDateFormatter(locator)
    ax_distance.xaxis.set_major_locator(locator)
    ax_distance.xaxis.set_major_formatter(formatter)
    fig.autofmt_xdate(rotation=30, ha="right")

    y_positions = list(range(len(rows)))
    ax_linear.hlines(y_positions, xmin=0.0, xmax=distances, color="0.82", linewidth=1.0)
    ax_linear.scatter(distances, y_positions, s=24, color="tab:orange", alpha=0.8)
    ax_linear.scatter([0.0], [len(rows) / 2], s=140, color="tab:blue", edgecolors="black", linewidths=0.8)

    start_y = y_positions[0]
    end_y = y_positions[-1]
    min_y = distances.index(min_row.distance_km)
    max_y = distances.index(max_row.distance_km)
    ax_linear.scatter([rows[0].distance_km], [start_y], s=55, color="green")
    ax_linear.scatter([rows[-1].distance_km], [end_y], s=55, color="red")
    ax_linear.scatter([min_row.distance_km], [min_y], s=70, facecolors="none", edgecolors="black", linewidths=1.1)
    ax_linear.scatter([max_row.distance_km], [max_y], s=70, facecolors="none", edgecolors="black", linewidths=1.1)

    def annotate_linear(label: str, x_value: float, y_value: float, *, x_offset: int = 8, y_offset: int = 0) -> None:
        ax_linear.annotate(
            label,
            (x_value, y_value),
            xytext=(x_offset, y_offset),
            textcoords="offset points",
            fontsize=8,
            va="center",
            bbox={"boxstyle": "round,pad=0.25", "facecolor": "white", "alpha": 0.92, "edgecolor": "0.7"},
            arrowprops={"arrowstyle": "-", "color": "0.4", "lw": 0.8},
        )

    annotate_linear("Earth", 0.0, len(rows) / 2, x_offset=10, y_offset=14)
    annotate_linear(f"Start\n{rows[0].day.isoformat()}", rows[0].distance_km, start_y, y_offset=14)
    annotate_linear(f"End\n{rows[-1].day.isoformat()}", rows[-1].distance_km, end_y, y_offset=-14)
    annotate_linear(f"Min\n{min_row.day.isoformat()}", min_row.distance_km, min_y, x_offset=-70, y_offset=-14)
    annotate_linear(f"Max\n{max_row.day.isoformat()}", max_row.distance_km, max_y, x_offset=-70, y_offset=14)

    tick_step = max(1, len(rows) // 8)
    ax_linear.set_yticks(y_positions[::tick_step])
    ax_linear.set_yticklabels([rows[index].day.isoformat() for index in y_positions[::tick_step]], fontsize=8)
    ax_linear.set_title("Linear View: Earth Fixed, Moon Distance by Day")
    ax_linear.set_xlabel("Distance from Earth (km)")
    ax_linear.set_ylabel("Date")
    ax_linear.grid(axis="x", alpha=0.35)
    ax_linear.set_xlim(left=-max(distances) * 0.05)
    ax_linear.invert_yaxis()

    summary_text = (
        f"Range: {rows[0].day.isoformat()} to {rows[-1].day.isoformat()}\n"
        f"UTC time: {utc_time.strftime('%H:%M')}\n"
        f"Minimum: {min_row.day.isoformat()} | {min_row.distance_km:,.1f} km\n"
        f"Maximum: {max_row.day.isoformat()} | {max_row.distance_km:,.1f} km"
    )
    fig.text(
        0.5,
        0.01,
        summary_text,
        ha="center",
        va="bottom",
        fontsize=9,
        bbox={"boxstyle": "round,pad=0.4", "facecolor": "white", "alpha": 0.9, "edgecolor": "0.75"},
    )

    fig.tight_layout(rect=(0, 0.08, 1, 1))
    fig.savefig(plot_output, dpi=150)
    plt.close(fig)


def print_summary(rows: list[DistanceRow], output_path: Path) -> None:
    min_row = min(rows, key=lambda item: item.distance_km)
    max_row = max(rows, key=lambda item: item.distance_km)

    print(f"Saved {len(rows)} rows to: {output_path}")
    print(f"Closest day : {min_row.day} -> {min_row.distance_km:,.1f} km")
    print(f"Farthest day: {max_row.day} -> {max_row.distance_km:,.1f} km")


class DistanceGui:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Earth-Moon Distance Daily")
        self.root.geometry("760x420")

        self.start_date = tk.StringVar(value=date.today().isoformat())
        self.days = tk.StringVar(value="30")
        self.time_utc = tk.StringVar(value="00:00")
        self.output_csv = tk.StringVar(value=str((Path.cwd() / "earth_moon_distance.csv").resolve()))
        self.plot_enabled = tk.BooleanVar(value=True)
        self.plot_output = tk.StringVar(value=str((Path.cwd() / "earth_moon_distance.png").resolve()))
        self.plot_tick_density = tk.IntVar(value=10)
        self.append_daily = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        for column in range(3):
            main.columnconfigure(column, weight=1)

        ttk.Label(main, text="Start date (YYYY-MM-DD)").grid(row=0, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.start_date).grid(row=1, column=0, sticky="ew", padx=(0, 8))

        ttk.Label(main, text="Days").grid(row=0, column=1, sticky="w")
        ttk.Entry(main, textvariable=self.days).grid(row=1, column=1, sticky="ew", padx=(0, 8))

        ttk.Label(main, text="UTC time (HH:MM)").grid(row=0, column=2, sticky="w")
        ttk.Entry(main, textvariable=self.time_utc).grid(row=1, column=2, sticky="ew")

        ttk.Label(main, text="CSV output").grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Entry(main, textvariable=self.output_csv).grid(row=3, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        ttk.Button(main, text="Browse", command=self._pick_csv).grid(row=3, column=2, sticky="ew")

        ttk.Checkbutton(main, text="Create PNG plot", variable=self.plot_enabled).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )
        ttk.Entry(main, textvariable=self.plot_output).grid(row=5, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        ttk.Button(main, text="Browse", command=self._pick_plot).grid(row=5, column=2, sticky="ew")

        ttk.Label(main, text="Date label density on PNG (fewer <-> more)").grid(
            row=6, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )
        tk.Scale(
            main,
            from_=4,
            to=20,
            orient="horizontal",
            variable=self.plot_tick_density,
            resolution=1,
            showvalue=True,
        ).grid(row=7, column=0, columnspan=3, sticky="ew")

        ttk.Checkbutton(
            main,
            text="Append one row for the selected day (skip duplicates)",
            variable=self.append_daily,
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(12, 0))

        ttk.Button(main, text="Run", command=self._run).grid(row=9, column=0, columnspan=3, sticky="ew", pady=(14, 0))

        actions = ttk.Frame(main)
        actions.grid(row=10, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        ttk.Button(actions, text="Open CSV", command=self._open_csv).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(actions, text="Open PNG", command=self._open_plot).grid(row=0, column=1, sticky="ew", padx=(4, 0))

        self.status_text = tk.Text(main, height=11, wrap="word")
        self.status_text.grid(row=11, column=0, columnspan=3, sticky="nsew", pady=(12, 0))
        main.rowconfigure(11, weight=1)

    def _pick_csv(self) -> None:
        picked = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if picked:
            self.output_csv.set(picked)

    def _pick_plot(self) -> None:
        picked = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")])
        if picked:
            self.plot_output.set(picked)

    def _log(self, message: str) -> None:
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)

    def _open_path(self, path: Path) -> None:
        if not path.exists():
            messagebox.showerror("File not found", f"File not found:\n{path}")
            return

        if os.name == "posix" and sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
            return

        # Fallback for non-macOS environments.
        try:
            if os.name == "nt":
                os.startfile(str(path))  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Open failed", str(exc))

    def _open_csv(self) -> None:
        self._open_path(Path(self.output_csv.get().strip()).expanduser().resolve())

    def _open_plot(self) -> None:
        plot_path_text = self.plot_output.get().strip()
        csv_path = Path(self.output_csv.get().strip()).expanduser().resolve()
        plot_path = (
            Path(plot_path_text).expanduser().resolve()
            if plot_path_text
            else csv_path.with_suffix(".png")
        )
        self._open_path(plot_path)

    def _run(self) -> None:
        self.status_text.delete("1.0", tk.END)
        try:
            start_day = parse_iso_date(self.start_date.get().strip())
            days = int(self.days.get().strip())
            utc_time = parse_utc_time(self.time_utc.get().strip())
            output_path = Path(self.output_csv.get().strip()).expanduser().resolve()

            if self.append_daily.get():
                rows = measure_distances(start_day=start_day, days=1, utc_time=utc_time)
                appended = append_csv(rows, output_path=output_path)
                self._log(f"CSV: {output_path}")
                self._log(f"Rows appended: {appended}")
            else:
                rows = measure_distances(start_day=start_day, days=days, utc_time=utc_time)
                write_csv(rows=rows, output_path=output_path)
                self._log(f"CSV overwritten: {output_path}")
                self._log(f"Rows saved: {len(rows)}")

            if self.plot_enabled.get():
                plot_path_text = self.plot_output.get().strip()
                plot_path = (
                    Path(plot_path_text).expanduser().resolve()
                    if plot_path_text
                    else output_path.with_suffix(".png")
                )
                create_plot(rows, plot_output=plot_path, utc_time=utc_time, max_ticks=self.plot_tick_density.get())
                self._log(f"Plot saved: {plot_path}")

            min_row = min(rows, key=lambda item: item.distance_km)
            max_row = max(rows, key=lambda item: item.distance_km)
            self._log(f"Closest day : {min_row.day} -> {min_row.distance_km:,.1f} km")
            self._log(f"Farthest day: {max_row.day} -> {max_row.distance_km:,.1f} km")
            messagebox.showinfo("Done", "Calculation completed successfully.")
        except Exception as exc:  # noqa: BLE001
            self._log(f"Error: {exc}")
            messagebox.showerror("Error", str(exc))

    def run(self) -> None:
        self.root.mainloop()


def main() -> int:
    args = parse_args()

    if args.gui or len(sys.argv) == 1:
        DistanceGui().run()
        return 0

    try:
        start_day = parse_iso_date(args.start_date)
        utc_time = parse_utc_time(args.time_utc)
        output_path = Path(args.output).expanduser().resolve()
        if args.append_daily:
            rows = measure_distances(start_day=start_day, days=1, utc_time=utc_time)
            appended = append_csv(rows=rows, output_path=output_path)
            print(f"CSV path: {output_path}")
            print(f"Rows appended: {appended}")
        else:
            rows = measure_distances(start_day=start_day, days=args.days, utc_time=utc_time)
            write_csv(rows=rows, output_path=output_path)
            print_summary(rows=rows, output_path=output_path)

        if args.plot:
            plot_path = (
                Path(args.plot_output).expanduser().resolve()
                if args.plot_output.strip()
                else output_path.with_suffix(".png")
            )
            create_plot(rows=rows, plot_output=plot_path, utc_time=utc_time)
            print(f"Plot saved: {plot_path}")
    except (argparse.ArgumentTypeError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 2
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(main())