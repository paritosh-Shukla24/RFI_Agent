"""
Configuration settings and environment validation
"""

import os
from dotenv import load_dotenv

# Debug mode for additional logging
DEBUG_MODE = False

def validate_environment():
    """Validate required environment variables are set"""
    # Load .env file if it exists
    load_dotenv()

    # Check for required API keys
    anthropic_key = os.getenv('ANTHROPIC_API_KEY')
    gemini_key = os.getenv('GEMINI_API_KEY')

    missing_keys = []

    if not anthropic_key:
        missing_keys.append('ANTHROPIC_API_KEY')

    if not gemini_key:
        missing_keys.append('GEMINI_API_KEY')

    if missing_keys:
        print("⚠️  Missing required environment variables:")
        for key in missing_keys:
            print(f"   - {key}")
        print("\n💡 Create a .env file with these variables or set them in your environment.")
        print("Example .env file content:")
        print("ANTHROPIC_API_KEY=your_anthropic_key_here")
        print("GEMINI_API_KEY=your_gemini_key_here")

        # Only raise error if both keys are missing - allow operation with just one LLM
        if len(missing_keys) > 1:
            raise ValueError("At least one LLM API key (ANTHROPIC_API_KEY or GEMINI_API_KEY) is required")
        else:
            print("\n⚠️  Continuing with limited functionality using available API key...")
    
    return True