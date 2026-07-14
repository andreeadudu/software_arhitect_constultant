"""
Application configuration.

This module contains all configurable settings used by the AI agent.

Future exercises may extend this file with:
- Model configuration
- API credentials
- Prompt templates
- Embedding settings
- Logging configuration
"""

# Numele modelului expus de serverul local (ex: Ollama, LM Studio etc.)
MODEL_NAME = "qwen3:1.7b"

# Adresa la care serverul local expune API-ul de chat
MODEL_ENDPOINT = "http://localhost:11434/api/chat"

# System prompt-ul agentului
SYSTEM_PROMPT = (
    "You are an experienced software architect, specialized in designing "
    "scalable, maintainable, and robust software systems. "
    "You provide clear, well-reasoned advice on software architecture, "
    "technology selection, design patterns, code structuring, "
    "technical trade-offs, and development best practices. "
    "You explain architectural decisions taking into account scalability, "
    "maintainability, testability, and cost. "
    "When you have tools available, you use them to better respond to "
    "the user's requests."

)