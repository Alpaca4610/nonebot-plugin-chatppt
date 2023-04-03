from pydantic import Extra, BaseModel
from typing import Optional


class Config(BaseModel, extra=Extra.ignore):
    openai_api_key: Optional[str] = ""
    openai_model_name: Optional[str] = "gpt-3.5-turbo"
    openai_http_proxy: Optional[str] = None
    slides_limit: Optional[str] = 10


class ConfigError(Exception):
    pass
