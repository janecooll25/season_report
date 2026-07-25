from __future__ import annotations

from pathlib import Path
from typing import Any

from docxtpl import DocxTemplate


def render_docx(template_path: str | Path, context: dict[str, Any], output_path: str | Path) -> Path:
    """Заполняет docx-шаблон значениями из context и сохраняет результат."""
    template_path = Path(template_path)
    output_path = Path(output_path)
    if not template_path.exists():
        raise FileNotFoundError(f"Шаблон не найден: {template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = DocxTemplate(str(template_path))
    doc.render(context)
    doc.save(str(output_path))
    return output_path
