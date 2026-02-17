import json
import time
from rich.console import Console
from rich.panel import Panel
from rich.live import Live
from rich.text import Text
from rich.table import Table
# MAX_TOOL_RESULT_LENGTH removed - no limits on tool results
from .base import get_current_time
from .file_tools import (
    create_file, read_file, update_file, delete_file,
    create_directory, list_directory, rename_file,
    copy_file, move_file, delete_directory
)
from .web_tools import (
    raw_web_browser, perform_searchapi_search,  search_scira, get_ip_geolocation
)
from .weather_tools import call_weather_api
from .terminal_tools import run_terminal_command
from .office_tools import generate_word_document, generate_excel_document
from .ppt_tools import generate_pptx_presentation

console = Console()

def _build_tool_panel(tool_name, spinner_char=None, is_complete=False):
    """Build a rich Panel for tool status display."""
    from rich.box import ROUNDED

    if is_complete:
        tbl = Table(show_header=False, show_edge=False, box=None, padding=(0, 1), expand=True)
        tbl.add_column("left", ratio=1)
        tbl.add_column("right", justify="right", width=12)
        tbl.add_row(
            Text.from_markup(f"[bold cyan]{tool_name}[/bold cyan]"),
            Text.from_markup("[bold green]complete[/bold green]")
        )
        return Panel(tbl, border_style="blue", box=ROUNDED, title="[bold]Tool[/bold]", title_align="left")
    else:
        tbl = Table(show_header=False, show_edge=False, box=None, padding=(0, 1), expand=True)
        tbl.add_column("left", ratio=1)
        tbl.add_column("right", justify="right", width=12)
        tbl.add_row(
            Text.from_markup(f"[bold cyan]{tool_name}[/bold cyan]"),
            Text.from_markup(f"[bold yellow]{spinner_char}[/bold yellow]")
        )
        return Panel(tbl, border_style="green", box=ROUNDED, title="[bold]Tool[/bold]", title_align="left")


def _run_spinner_animation(tool_name, stop_event, live_obj):
    """Animation thread: only shows spinning frames, stops when event is set."""
    spinner_frames = ["─", "╲", "│", "╱"]
    idx = 0
    while not stop_event.is_set():
        frame = spinner_frames[idx % len(spinner_frames)]
        panel = _build_tool_panel(tool_name, spinner_char=frame, is_complete=False)
        live_obj.update(panel)
        idx += 1
        time.sleep(0.15)


def execute_tool_call(tool_name, tool_params):
    """Execute tool call with dynamic spinning line status display.

    Flow:
    1. Start Live(transient=True) — all intermediate frames are erased on exit
    2. Animation thread shows spinning ─ ╲ │ ╱ in the panel
    3. Tool executes (or for terminal: brief spinner then stop)
    4. Stop animation, exit Live (transient clears everything)
    5. Print static complete panel (no blank line for tiling)
    6. (For terminal tools: run command output after the panel)
    """
    import threading

    stop_event = threading.Event()

    try:
        if tool_name == "run_terminal_command":
            # Show brief spinner, then complete panel, then run command
            with Live(refresh_per_second=10, transient=True) as live:
                anim_thread = threading.Thread(
                    target=_run_spinner_animation,
                    args=(tool_name, stop_event, live)
                )
                anim_thread.start()
                time.sleep(0.6)
                stop_event.set()
                anim_thread.join()

            # Static complete panel (no extra blank line for tiling)
            console.print(_build_tool_panel(tool_name, is_complete=True))

            result = run_terminal_command(**tool_params)

            console.print()  # Blank line after terminal output
            return result

        # All other tools: spinner runs during execution
        with Live(refresh_per_second=10, transient=True) as live:
            anim_thread = threading.Thread(
                target=_run_spinner_animation,
                args=(tool_name, stop_event, live)
            )
            anim_thread.start()

            result = _dispatch_tool(tool_name, tool_params)

            # Stop animation
            stop_event.set()
            anim_thread.join()

        # Static complete panel (Live transient already cleared animation, no extra blank line for tiling)
        console.print(_build_tool_panel(tool_name, is_complete=True))

        return result

    except Exception as e:
        stop_event.set()
        return {"error": f"Tool execution error: {str(e)}"}


def _dispatch_tool(tool_name, tool_params):
    """Dispatch tool call to the appropriate function."""
    dispatch_map = {
        "raw_web_browser": lambda p: raw_web_browser(**p),
        "perform_searchapi_search": lambda p: perform_searchapi_search(**p),
        "perform_bing_search": lambda p: perform_bing_search(**p),
        "perform_google_search": lambda p: perform_google_search(**p),
        "search_scira": lambda p: search_scira(**p),
        "get_ip_geolocation": lambda p: get_ip_geolocation(**p),
        "generate_word_document": lambda p: generate_word_document(**p),
        "generate_excel_document": lambda p: generate_excel_document(**p),
        "generate_pptx_presentation": lambda p: generate_pptx_presentation(**p),
        "create_file": lambda p: create_file(**p),
        "read_file": lambda p: read_file(**p),
        "update_file": lambda p: update_file(**p),
        "delete_file": lambda p: delete_file(**p),
        "create_directory": lambda p: create_directory(**p),
        "list_directory": lambda p: list_directory(**p),
        "rename_file": lambda p: rename_file(**p),
        "copy_file": lambda p: copy_file(**p),
        "move_file": lambda p: move_file(**p),
        "delete_directory": lambda p: delete_directory(**p),
        "get_current_time": lambda p: get_current_time(),
    }

    # Special handling for weather API
    if tool_name == "call_weather_api":
        api_type = tool_params.pop("api_type", "current")
        return call_weather_api(api_type, **tool_params)

    handler = dispatch_map.get(tool_name)
    if handler:
        return handler(tool_params)
    return {"error": f"Unknown tool: {tool_name}"}


def compress_tool_result(tool_name, result):
    """Compress tool result to avoid context being too long"""
    if isinstance(result, dict) and "error" in result:
        return result

    # No truncation - return full result
    return result
