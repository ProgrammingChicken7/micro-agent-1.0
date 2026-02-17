import os
import json

# --- Configuration Files ---
MODEL_CONFIG_FILE = "model_config.json"
WORKSPACE_CONFIG_FILE = "workspace_config.json"
TOOLS_CONFIG_FILE = "tools_config.json"

# --- Context Management Configuration (Dynamic) ---
# These will be set dynamically based on model's context_limit
MAX_HISTORY_LENGTH = 40          # Max conversation turns to keep in full
# MAX_TOOL_RESULT_LENGTH removed - no limits on tool results
MAX_CONTEXT_TOKENS_ESTIMATE = 28000  # Estimated token budget (conservative for most models)
CHARS_PER_TOKEN_ESTIMATE = 3.5   # Average chars per token (conservative for mixed CJK/English)
SUMMARY_TRIGGER_RATIO = 0.7      # Trigger summarization when context reaches 70% of budget

# --- Context Limit Presets ---
CONTEXT_LIMIT_PRESETS = {
    "32K": {"tokens": 32000, "history_length": 40},
    "64K": {"tokens": 64000, "history_length": 60},
    "128K": {"tokens": 128000, "history_length": 80},
    "192K": {"tokens": 192000, "history_length": 100},
    "256K": {"tokens": 256000, "history_length": 120},
    "512K": {"tokens": 512000, "history_length": 180},
    "1M": {"tokens": 1000000, "history_length": 300},
    "2M": {"tokens": 2000000, "history_length": 500},
}

def load_model_config():
    """Load model configuration from file"""
    if os.path.exists(MODEL_CONFIG_FILE):
        try:
            with open(MODEL_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

def save_model_config(config):
    """Save model configuration to file"""
    with open(MODEL_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

def get_context_params(context_limit):
    """Get context management parameters based on context limit.
    
    Args:
        context_limit: String like "32K", "128K", "1M", or custom like "150K"
    
    Returns:
        dict with keys: max_tokens, trigger_tokens, max_history
    """
    # Check if it's a preset
    if context_limit in CONTEXT_LIMIT_PRESETS:
        preset = CONTEXT_LIMIT_PRESETS[context_limit]
        max_tokens = preset["tokens"]
        history_length = preset["history_length"]
    else:
        # Parse custom limit (e.g., "150K", "150", "1.5M")
        try:
            if context_limit.upper().endswith('K'):
                max_tokens = int(float(context_limit[:-1]) * 1000)
            elif context_limit.upper().endswith('M'):
                max_tokens = int(float(context_limit[:-1]) * 1000000)
            else:
                max_tokens = int(float(context_limit) * 1000)  # Assume input is in K tokens
            
            # Calculate adaptive parameters based on token count
            if max_tokens <= 32000:
                history_length = 40
            elif max_tokens <= 64000:
                history_length = 60
            elif max_tokens <= 128000:
                history_length = 80
            elif max_tokens <= 256000:
                history_length = 120
            elif max_tokens <= 512000:
                history_length = 180
            elif max_tokens <= 1000000:
                history_length = 300
            else:
                history_length = 500
        except (ValueError, AttributeError):
            # Fallback to 64K if parsing fails
            preset = CONTEXT_LIMIT_PRESETS["64K"]
            max_tokens = preset["tokens"]
            history_length = preset["history_length"]
    
    return {
        "max_tokens": int(max_tokens * 0.85),  # 85% of limit as safe budget
        "trigger_tokens": int(max_tokens * 0.70),  # 70% triggers compression
        "max_history": history_length
    }

def apply_context_params(context_limit):
    """Apply context parameters globally based on context limit."""
    global MAX_HISTORY_LENGTH, MAX_CONTEXT_TOKENS_ESTIMATE
    
    params = get_context_params(context_limit)
    MAX_HISTORY_LENGTH = params["max_history"]
    MAX_CONTEXT_TOKENS_ESTIMATE = params["max_tokens"]

def load_tools_config():
    """Load tool API keys from config file"""
    if os.path.exists(TOOLS_CONFIG_FILE):
        try:
            with open(TOOLS_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def get_tool_config(key, default=""):
    """Get a specific tool configuration"""
    config = load_tools_config()
    return config.get(key, default)

def estimate_tokens(text):
    """Estimate token count from text length."""
    if not text:
        return 0
    return int(len(str(text)) / CHARS_PER_TOKEN_ESTIMATE)

def estimate_messages_tokens(messages):
    """Estimate total token count for a list of messages."""
    total = 0
    for msg in messages:
        total += 4  # message overhead
        content = msg.get("content", "")
        if content:
            total += estimate_tokens(content)
        # Tool calls in assistant messages
        if "tool_calls" in msg:
            for tc in msg["tool_calls"]:
                total += estimate_tokens(json.dumps(tc))
    return total
