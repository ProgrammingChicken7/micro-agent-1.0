import openai
import traceback
from rich.console import Console
from rich.markdown import Markdown
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.table import Table
from rich.text import Text
from rich.live import Live
import json
import os
import sys
import requests
import time
import threading

# Import configurations and tools
from config import (
    MAX_HISTORY_LENGTH, MODEL_CONFIG_FILE, WORKSPACE_CONFIG_FILE, TOOLS_CONFIG_FILE,
    SUMMARY_TRIGGER_RATIO, estimate_messages_tokens
)
from tools.base import load_workspace_config, get_workspace_path
from tools.manager import execute_tool_call, compress_tool_result
from tools.definitions import TOOLS

# --- Global Variables ---
console = Console()
history = []

# --- System Prompt ---
SYSTEM_PROMPT = {
    "role": "system",
    "content": """You are a powerful AI assistant with access to various tools. You can create professional documents, search the web, check weather, run terminal commands, and manage files.

IMPORTANT GUIDELINES FOR DOCUMENT GENERATION:
1. When asked to create Word documents, ALWAYS use generate_word_document with rich styling:
   - Start formal documents with cover_page and toc
   - Use charts (type:"chart") to visualize any data — prefer charts over plain tables
   - Use consistent color schemes (e.g. headings in "#003366", tables with "#4472C4" headers)
   - Add page_numbers, headers, and footers for professional appearance
   - Use rich_paragraph for mixed formatting, quote blocks for key insights
   - Set document_settings with appropriate fonts, margins, and line_spacing

2. When asked to create Excel documents, ALWAYS use generate_excel_document with:
   - Professional header_style with colored backgrounds
   - stripe_colors for alternating row colors
   - freeze_panes to keep headers visible
   - auto_filter for data tables
   - Charts to visualize key metrics
   - conditional_formatting to highlight important values

3. When asked to create PowerPoint presentations, ALWAYS use generate_pptx_presentation with:
   - Choose an appropriate theme: "ocean", "forest", "sunset", "royal", "midnight", "coral", "tech", "elegant"
   - ALWAYS start with slide_type:"title" and end with slide_type:"ending"
   - Use slide_type:"section" to divide major parts of the presentation
   - Use slide_type:"stats" for key numbers and metrics
   - Use slide_type:"chart" with native chart types (column, bar, line, pie, etc.) for data visualization
   - Use slide_type:"cards" or "three_column" for feature lists and comparisons
   - Use slide_type:"timeline" for processes, roadmaps, and historical events
   - Use slide_type:"comparison" for pros/cons or before/after comparisons
   - Use slide_type:"table" for structured data display
   - Use slide_type:"quote" for impactful statements
   - Use slide_type:"two_column" for side-by-side content
   - Vary slide types for visual interest — don't use the same type repeatedly
   - Use bullets for lists, paragraphs for detailed text
   - Generate 8-15 slides for a comprehensive presentation

4. For complex documents, break the work into multiple tool calls if needed.
5. Always use the workspace for file operations.
6. When generating charts in Word, use type:"chart" blocks with appropriate chart_type.
"""
}

# --- Rainbow Logo Functions ---
def rgb_to_ansi(r, g, b):
    if r == g == b:
        if r < 8: return 16
        if r > 248: return 231
        return round(((r - 8) / 247) * 24) + 232
    return 16 + (36 * round(r / 255 * 5) + 6 * round(g / 255 * 5) + round(b / 255 * 5))

def get_rainbow_color_by_line(line_index, char_index, total_lines, line_length):
    rainbow_colors = [(255, 0, 0), (255, 127, 0), (255, 255, 0), (0, 255, 0), (0, 0, 255), (75, 0, 130), (148, 0, 211)]
    color_index = int((line_index / total_lines) * len(rainbow_colors))
    if color_index >= len(rainbow_colors): color_index = len(rainbow_colors) - 1
    if line_length > 0:
        next_color_index = min(color_index + 1, len(rainbow_colors) - 1)
        progress = char_index / line_length
        color1, color2 = rainbow_colors[color_index], rainbow_colors[next_color_index]
        r = int(color1[0] + (color2[0] - color1[0]) * progress)
        g = int(color1[1] + (color2[1] - color1[1]) * progress)
        b = int(color1[2] + (color2[2] - color1[2]) * progress)
        return (r, g, b)
    return rainbow_colors[color_index]

def print_rainbow_ascii(ascii_art):
    lines = ascii_art.split('\n')
    total_lines = len([line for line in lines if line.strip()])
    line_index = 0
    for line in lines:
        if line.strip():
            colored_line, line_length = "", len(line)
            for char_index, char in enumerate(line):
                if char.strip():
                    r, g, b = get_rainbow_color_by_line(line_index, char_index, total_lines, line_length)
                    ansi_code = rgb_to_ansi(r, g, b)
                    colored_line += f"\033[38;5;{ansi_code}m{char}"
                else: colored_line += char
            colored_line += "\033[0m"
            print(colored_line)
            line_index += 1
        else: print(line)

# --- Model Management ---
def load_model_config():
    if os.path.exists(MODEL_CONFIG_FILE):
        try:
            with open(MODEL_CONFIG_FILE, 'r', encoding='utf-8') as f: return json.load(f)
        except: return []
    return []

def save_model_config(models):
    with open(MODEL_CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(models, f, indent=2, ensure_ascii=False)

def fetch_available_models(api_base, api_key):
    try:
        url = api_base.rstrip('/') + '/models'
        headers = {"Authorization": f"Bearer {api_key}"}
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if isinstance(data, dict) and 'data' in data: return [m['id'] for m in data['data']]
            elif isinstance(data, list): return [m['id'] for m in data]
        return []
    except: return []

def setup_models():
    from config import CONTEXT_LIMIT_PRESETS
    console.print("[bold cyan]First-time setup: Configure AI models[/bold cyan]")
    models = load_model_config()
    while True:
        api_base = console.input("[green]API base URL: [/green]").strip()
        if not api_base: break
        api_key = console.input("[green]API key: [/green]").strip()
        if not api_key: continue
        available_ids = fetch_available_models(api_base, api_key)
        if available_ids:
            choice = console.input("[yellow]Add all (a), select (s), or manual (m)? [a/s/m]: [/yellow]").lower()
            if choice == 'a':
                # Ask for context limit once for batch add
                console.print("[cyan]Select context limit for all models:[/cyan]")
                for i, limit in enumerate(["32K", "64K", "128K", "192K", "256K", "512K", "1M", "2M"], 1):
                    console.print(f"  {i}. {limit}")
                console.print("  9. Custom (input in K tokens)")
                try:
                    limit_choice = console.input("[yellow]Choice (default 2): [/yellow]").strip() or "2"
                    if limit_choice == "9":
                        custom_k = console.input("[yellow]Enter context limit (K tokens, e.g., 150): [/yellow]").strip()
                        context_limit = f"{custom_k}K"
                    else:
                        context_limit = ["32K", "64K", "128K", "192K", "256K", "512K", "1M", "2M"][int(limit_choice) - 1]
                except:
                    context_limit = "64K"
                for mid in available_ids:
                    models.append({"display_name": mid, "name": mid, "api_base": api_base, "api_key": api_key, "context_limit": context_limit})
            elif choice == 's':
                for i, mid in enumerate(available_ids, 1): console.print(f"{i}. {mid}")
                indices = console.input("[yellow]Indices (e.g. 1,2): [/yellow]").split(',')
                for idx in indices:
                    try:
                        mid = available_ids[int(idx.strip()) - 1]
                        display_name = console.input(f"[green]Display name for {mid}: [/green]").strip() or mid
                        # Ask for context limit
                        console.print(f"[cyan]Context limit for {display_name}:[/cyan]")
                        for i, limit in enumerate(["32K", "64K", "128K", "192K", "256K", "512K", "1M", "2M"], 1):
                            console.print(f"  {i}. {limit}")
                        console.print("  9. Custom (input in K tokens)")
                        try:
                            limit_choice = console.input("[yellow]Choice (default 2): [/yellow]").strip() or "2"
                            if limit_choice == "9":
                                custom_k = console.input("[yellow]Enter context limit (K tokens, e.g., 150): [/yellow]").strip()
                                context_limit = f"{custom_k}K"
                            else:
                                context_limit = ["32K", "64K", "128K", "192K", "256K", "512K", "1M", "2M"][int(limit_choice) - 1]
                        except:
                            context_limit = "64K"
                        models.append({"display_name": display_name, "name": mid, "api_base": api_base, "api_key": api_key, "context_limit": context_limit})
                    except: continue
            else:
                mid = console.input("[green]Model ID: [/green]").strip()
                display_name = console.input("[green]Display name: [/green]").strip()
                # Ask for context limit
                console.print("[cyan]Select context limit:[/cyan]")
                for i, limit in enumerate(["32K", "64K", "128K", "192K", "256K", "512K", "1M", "2M"], 1):
                    console.print(f"  {i}. {limit}")
                console.print("  9. Custom (input in K tokens)")
                try:
                    limit_choice = console.input("[yellow]Choice (default 2): [/yellow]").strip() or "2"
                    if limit_choice == "9":
                        custom_k = console.input("[yellow]Enter context limit (K tokens, e.g., 150): [/yellow]").strip()
                        context_limit = f"{custom_k}K"
                    else:
                        context_limit = ["32K", "64K", "128K", "192K", "256K", "512K", "1M", "2M"][int(limit_choice) - 1]
                except:
                    context_limit = "64K"
                models.append({"display_name": display_name, "name": mid, "api_base": api_base, "api_key": api_key, "context_limit": context_limit})
        else:
            mid = console.input("[green]Model ID: [/green]").strip()
            display_name = console.input("[green]Display name: [/green]").strip()
            # Ask for context limit
            console.print("[cyan]Select context limit:[/cyan]")
            for i, limit in enumerate(["32K", "64K", "128K", "192K", "256K", "512K", "1M", "2M"], 1):
                console.print(f"  {i}. {limit}")
            console.print("  9. Custom (input in K tokens)")
            try:
                limit_choice = console.input("[yellow]Choice (default 2): [/yellow]").strip() or "2"
                if limit_choice == "9":
                    custom_k = console.input("[yellow]Enter context limit (K tokens, e.g., 150): [/yellow]").strip()
                    context_limit = f"{custom_k}K"
                else:
                    context_limit = ["32K", "64K", "128K", "192K", "256K", "512K", "1M", "2M"][int(limit_choice) - 1]
            except:
                context_limit = "64K"
            models.append({"display_name": display_name, "name": mid, "api_base": api_base, "api_key": api_key, "context_limit": context_limit})
        if console.input("[yellow]Add another? (y/n): [/yellow]").lower() != 'y': break
    if models: save_model_config(models); return True
    return False

def select_model():
    models = load_model_config()
    if not models: return None
    for i, m in enumerate(models, 1):
        context_limit = m.get('context_limit', '64K')
        console.print(f"[green]{i}. {m.get('display_name', m['name'])}[/green] [dim cyan]({context_limit})[/dim cyan]")
    while True:
        try:
            choice = int(console.input("[yellow]Select model: [/yellow]"))
            if 1 <= choice <= len(models):
                selected = models[choice - 1]
                # Ensure context_limit exists
                if 'context_limit' not in selected:
                    selected['context_limit'] = '64K'
                return selected
        except: pass

# --- Tool Config ---
def setup_tools_config():
    console.print("[bold cyan]First-time setup: Tool API Keys[/bold cyan]")
    config = {}
    tools = [("WEATHERAPI_KEY", "WeatherAPI"), ("SEARCHAPI_API_KEY", "SearchAPI"), ("SCIRA_API_KEY", "Scira"), ("IPGEOLOCATION_API_KEY", "IPGeo")]
    for key, desc in tools:
        val = console.input(f"[green]{desc} Key: [/green]").strip()
        if val: config[key] = val
    with open(TOOLS_CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(config, f, indent=2, ensure_ascii=False)
    return True

# --- Workspace ---
def select_workspace():
    """Select workspace directory — supports both GUI and terminal fallback."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        path = filedialog.askdirectory(title="Select Workspace")
        root.destroy()
    except:
        path = console.input("[green]Enter workspace directory path: [/green]").strip()

    if path:
        os.makedirs(path, exist_ok=True)
        with open(WORKSPACE_CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump({"path": path}, f, indent=2, ensure_ascii=False)
        return path
    return None


# ============================================================================
#  CONTEXT MANAGEMENT
# ============================================================================

def manage_context(messages):
    """
    Manage conversation context by removing oldest 30% when limit is reached.
    
    Strategy:
    - No compression or truncation
    - When context reaches limit, delete the oldest 30% of messages
    - Always keep system prompt intact
    - No restrictions on tool usage
    """
    if not messages:
        return messages

    import config
    estimated_tokens = estimate_messages_tokens(messages)
    token_budget = config.MAX_CONTEXT_TOKENS_ESTIMATE

    # Only act when we've exceeded the budget
    if estimated_tokens <= token_budget:
        return messages

    console.print(f"[dim yellow]Context limit reached: ~{estimated_tokens} tokens, removing oldest 30%...[/dim yellow]")

    # Separate system messages and conversation
    system_msgs = [m for m in messages if m.get("role") == "system"]
    conv_msgs = [m for m in messages if m.get("role") != "system"]

    if len(conv_msgs) <= 2:
        return messages  # Too few messages to remove

    # Calculate how many messages to remove (30% of conversation messages)
    remove_count = max(1, int(len(conv_msgs) * 0.3))
    
    # Keep the newest 70%
    kept_msgs = conv_msgs[remove_count:]
    
    # Reconstruct with system messages + kept conversation
    result = system_msgs + kept_msgs

    new_tokens = estimate_messages_tokens(result)
    console.print(f"[dim green]Context reduced: {estimated_tokens} → ~{new_tokens} tokens (removed {remove_count} messages)[/dim green]")
    return result


# ============================================================================    return result


# ============================================================================
#  AI LOGIC
# ============================================================================

def _clean_message_content(content):
    """Clean message content to remove surrogate characters that cause UTF-8 encoding errors."""
    if isinstance(content, str):
        # Remove surrogate characters (U+D800 to U+DFFF)
        return content.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
    return content

def _clean_messages(messages):
    """Clean all messages to ensure they can be safely encoded as UTF-8."""
    cleaned = []
    for msg in messages:
        cleaned_msg = {"role": msg["role"]}
        if "content" in msg:
            cleaned_msg["content"] = _clean_message_content(msg["content"])
        if "tool_calls" in msg:
            cleaned_msg["tool_calls"] = msg["tool_calls"]
        if "tool_call_id" in msg:
            cleaned_msg["tool_call_id"] = msg["tool_call_id"]
        if "reasoning_content" in msg:
            cleaned_msg["reasoning_content"] = _clean_message_content(msg["reasoning_content"])
        cleaned.append(cleaned_msg)
    return cleaned

def handle_llm_response(prompt):
    global history
    history.append({"role": "user", "content": prompt})
    selected_model = getattr(handle_llm_response, 'selected_model', None)
    if not selected_model: return

    client = openai.OpenAI(api_key=selected_model['api_key'], base_url=selected_model['api_base'])

    # Build messages with system prompt + managed context
    messages = [SYSTEM_PROMPT] + manage_context(history)
    # Clean messages to prevent UTF-8 encoding errors
    messages = _clean_messages(messages)

    while True:
        # Start a fresh status spinner for each LLM call
        with console.status("[bold cyan]AI thinking...[/bold cyan]", spinner="dots") as status:
            try:
                stream = client.chat.completions.create(
                    model=selected_model['name'],
                    messages=messages,
                    stream=True,
                    tools=TOOLS,
                    tool_choice="auto"
                )
            except openai.APIError as e:
                error_msg = str(e)
                if "context" in error_msg.lower() or "token" in error_msg.lower() or "length" in error_msg.lower():
                    # Context too long — aggressive compression and retry
                    console.print("[bold yellow]Context too long, compressing further...[/bold yellow]")
                    # Keep only last 4 messages
                    recent = history[-4:] if len(history) > 4 else history
                    messages = [SYSTEM_PROMPT, {
                        "role": "system",
                        "content": "[Previous conversation was too long and has been trimmed. Continuing from recent context.]"
                    }] + recent
                    # Clean messages to prevent UTF-8 encoding errors
                    messages = _clean_messages(messages)
                    try:
                        stream = client.chat.completions.create(
                            model=selected_model['name'],
                            messages=messages,
                            stream=True,
                            tools=TOOLS,
                            tool_choice="auto"
                        )
                    except Exception as e2:
                        console.print(f"[red]API Error after compression: {str(e2)}[/red]")
                        return
                else:
                    console.print(f"[red]API Error: {error_msg}[/red]")
                    return

            content_parts, tool_calls, live_content = [], [], ""
            reasoning_parts = []  # For thinking models like Kimi-K2.5
            for chunk in stream:
                if chunk.choices[0].delta.content:
                    delta = chunk.choices[0].delta.content
                    content_parts.append(delta); live_content += delta
                    status.update(f"[bold cyan]AI Response:[/bold cyan] {live_content[:100]}...")
                # Handle reasoning_content for thinking models
                if hasattr(chunk.choices[0].delta, 'reasoning_content') and chunk.choices[0].delta.reasoning_content:
                    reasoning_parts.append(chunk.choices[0].delta.reasoning_content)
                if chunk.choices[0].delta.tool_calls:
                    for tc_delta in chunk.choices[0].delta.tool_calls:
                        if tc_delta.index >= len(tool_calls):
                            tool_calls.append({"id": tc_delta.id, "type": "function", "function": {"name": tc_delta.function.name or "", "arguments": tc_delta.function.arguments or ""}})
                        else:
                            tc = tool_calls[tc_delta.index]
                            if tc_delta.function.name: tc["function"]["name"] = tc_delta.function.name
                            if tc_delta.function.arguments: tc["function"]["arguments"] += tc_delta.function.arguments

            assistant_content = "".join(content_parts)
            assistant_msg = {"role": "assistant", "content": assistant_content}
            # Add reasoning_content if present (for thinking models)
            if reasoning_parts:
                assistant_msg["reasoning_content"] = "".join(reasoning_parts)
            if tool_calls: assistant_msg["tool_calls"] = tool_calls
            messages.append(assistant_msg); history.append(assistant_msg)

        # --- Outside console.status context ---
        # Tool calls are processed outside the status spinner so panels display correctly
        if tool_calls:
            tool_results = []
            for tc in tool_calls:
                name, args_str = tc["function"]["name"], tc["function"]["arguments"]
                try:
                    args = json.loads(args_str)
                    res = execute_tool_call(name, args)
                    res = compress_tool_result(name, res)
                    tool_results.append({"tool_call_id": tc["id"], "role": "tool", "content": json.dumps(res, ensure_ascii=False)})
                except Exception as e:
                    tool_results.append({"tool_call_id": tc["id"], "role": "tool", "content": json.dumps({"error": str(e)})})

            messages.extend(tool_results); history.extend(tool_results)

            # Re-check context size before next iteration
            import config
            estimated = estimate_messages_tokens(messages)
            if estimated > config.MAX_CONTEXT_TOKENS_ESTIMATE * 0.8:
                messages = [SYSTEM_PROMPT] + manage_context(history)
                # Clean messages to prevent UTF-8 encoding errors
                messages = _clean_messages(messages)

            continue
        
        # Display context usage after response
        import config
        current_tokens = estimate_messages_tokens(messages)
        usage_percent = (current_tokens / config.MAX_CONTEXT_TOKENS_ESTIMATE) * 100
        
        # Color based on usage level
        if usage_percent < 50:
            color = "green"
        elif usage_percent < 70:
            color = "yellow"
        else:
            color = "red"
        
        console.print(f"[dim {color}]Context: {current_tokens:,}/{config.MAX_CONTEXT_TOKENS_ESTIMATE:,} tokens ({usage_percent:.1f}%)[/dim {color}]\n")
        break

    if assistant_content:
        console.print(Panel(Markdown(assistant_content), title="[bold cyan]AI Response[/bold cyan]", border_style="cyan", padding=(1, 2)))

def show_logo():
    logo = r'''
 ___  ___  ___  ________  ________  ________          ________  ________  _______   ________   _________   
|\  \|\  \|\  \|\   ____\|\   __  \|\   __  \        |\   __  \|\   ____\|\  ___ \ |\   ___  \|\___   ___\ 
\ \  \\\  \ \  \ \  \___|\ \  \|\  \ \  \|\  \       \ \  \|\  \ \  \___|\ \   __/|\ \  \\ \  \|___ \  \_| 
 \ \   __  \ \  \ \  \    \ \   _  _\ \  \\\  \       \ \   __  \ \  \  __\ \  \_|/_\ \  \\ \  \   \ \  \  
  \ \  \ \  \ \  \ \  \____\ \  \\  \\ \  \\\  \       \ \  \ \  \ \  \|\  \ \  \_|\ \ \  \\ \  \   \ \  \ 
   \ \__\ \__\ \__\ \_______\ \__\\ _\\ \_______\       \ \__\ \__\ \_______\ \_______\ \__\\ \__\   \ \__\
    \|__|\|__|\|__|\|_______|\|__|\|__|\|_______|        \|__|\|__|\|_______|\|_______|\|__| \|__|    \|__|
'''
    print_rainbow_ascii(logo)

def main():
    show_logo()
    if not os.path.exists(MODEL_CONFIG_FILE) or not load_model_config(): setup_models()
    if not os.path.exists(TOOLS_CONFIG_FILE): setup_tools_config()
    if not os.path.exists(WORKSPACE_CONFIG_FILE): select_workspace()

    selected_model = select_model()
    if not selected_model: return
    handle_llm_response.selected_model = selected_model
    
    # Apply context parameters based on model's context limit
    from config import apply_context_params, get_context_params
    context_limit = selected_model.get('context_limit', '64K')
    apply_context_params(context_limit)
    params = get_context_params(context_limit)
    
    console.print(f"\n[bold green]Model loaded: {selected_model.get('display_name', selected_model['name'])}[/bold green]")
    console.print(f"[dim cyan]Context limit: {context_limit} (Budget: {params['max_tokens']:,} tokens, History: {params['max_history']} turns)[/dim cyan]")
    console.print("[dim]Type 'exit' or 'quit' to exit. Context is managed automatically.[/dim]\n")

    while True:
        try:
            user_input = console.input("[bold blue]You: [/bold blue]").strip()
            if user_input.lower() in ['exit', 'quit']: break
            if user_input: handle_llm_response(user_input)
        except KeyboardInterrupt: break
        except Exception as e:
            console.print(f"[red]Error: {str(e)}[/red]")
            console.print(f"[dim red]{traceback.format_exc()}[/dim red]")

if __name__ == "__main__": main()
