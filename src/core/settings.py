from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    path_bar: str
    path_src: str
    path_sitegroup: str
    path_plan: str
    path_report_ag: str

    model_config = SettingsConfigDict(
        env_file=Path(r"D:\Code\ai-agents-data-portal\gold-promo-agents\.env"),
        env_file_encoding="utf-8",
    )


settings = Settings()
