def fence_md(txt: str) -> str:
    "```md\n...\n``` 包装，防 prompt 注入"
    return f"```md\n{txt}\n```"