import os
from pathlib import Path
from dotenv import load_dotenv

class Config:
    def __init__(self):
        self.config_dir = Path.home() / ".excel-gpt-translator"
        self.config_file = self.config_dir / "config.env"
        self._ensure_config_dir()
        load_dotenv(self.config_file)
        
        # Default settings
        self.api_key = os.getenv("OPENAI_API_KEY", "")
        self.default_languages = [
            "English", "Spanish", "French", "German", "Chinese",
            "Japanese", "Korean", "Russian", "Arabic", "Portuguese"
        ]
        self.default_prompt = (
            "As a professional translator, please translate the following content from {current_lang} to {target_lang}. "
            "Maintain the original meaning, tone, and context while ensuring the translation is culturally appropriate. "
            "{field_context}"
            "For technical terms, use industry-standard terminology. "
            "For business content, maintain formal language and professional tone. "
            "Text to translate:\n\n{text}"
        )
    
    def _ensure_config_dir(self):
        """Ensure the configuration directory exists."""
        self.config_dir.mkdir(exist_ok=True)
        if not self.config_file.exists():
            self.config_file.touch()
    
    def save_api_key(self, api_key: str):
        """Save the OpenAI API key to the configuration file."""
        with open(self.config_file, "w") as f:
            f.write(f"OPENAI_API_KEY={api_key}\n")
        self.api_key = api_key
    
    def get_api_key(self) -> str:
        """Get the OpenAI API key."""
        return self.api_key
    
    def get_supported_languages(self) -> list:
        """Get the list of supported languages."""
        return self.default_languages
    
    def format_prompt(self, current_lang: str, target_lang: str, text: str, field: str = "") -> str:
        """Format the translation prompt with optional field/industry context."""
        field_context = f"This content is specifically related to the {field} field/industry. " if field else ""
        return self.default_prompt.format(
            current_lang=current_lang,
            target_lang=target_lang,
            text=text,
            field_context=field_context
        )
    
    def get_default_prompt(self) -> str:
        """Get the default translation prompt."""
        return self.default_prompt 