import subprocess
import sys
import os
from rich.console import Console
from .base import get_workspace_path

console = Console()

def run_terminal_command(command):
    """
    Run a terminal command and stream output to the user.
    """
    workspace = get_workspace_path()
    cwd = workspace if workspace and os.path.exists(workspace) else os.getcwd()

    try:
        process = subprocess.Popen(
            command,
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=cwd,
            bufsize=1,
            universal_newlines=True
        )

        full_output = []
        console.print(f"[bold blue]Executing: {command}[/bold blue]")

        for line in process.stdout:
            print(line, end="", flush=True)
            full_output.append(line)

        process.wait()

        # Ensure output ends with a newline
        output_text = "".join(full_output)
        if output_text and not output_text.endswith('\n'):
            print()

        return {
            "success": process.returncode == 0,
            "exit_code": process.returncode,
            "output": output_text
        }
    except Exception as e:
        return {"error": str(e)}
