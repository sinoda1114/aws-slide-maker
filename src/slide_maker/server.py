"""MCP server — exposes slide-generation tools to Cursor / Claude Code."""
from __future__ import annotations

import os
from pathlib import Path

import yaml
from mcp.server.fastmcp import FastMCP

from .generator import generate
from .planner import parse_input, spec_from_description, spec_from_yaml
from .schema import PresentationSpec

mcp = FastMCP(
    "slide-maker",
    instructions=(
        "AI-powered PowerPoint (PPTX) generator. "
        "Use `create_presentation` for natural-language prompts or YAML specs. "
        "Use `preview_spec` to inspect the generated YAML before building. "
        "Set ANTHROPIC_API_KEY in your environment before calling."
    ),
)

# ---------------------------------------------------------------------------
# Tool: create_presentation
# ---------------------------------------------------------------------------

@mcp.tool()
def create_presentation(
    description: str,
    output_path: str = "./output/presentation.pptx",
    theme: str = "aws",
    model: str = "claude-sonnet-4-6",
) -> str:
    """
    PowerPointプレゼンテーション（PPTX）を生成します。

    Args:
        description:  自然言語での説明（例: "AWSセキュリティのベストプラクティス 10枚"）
                      またはYAML形式のスライド仕様書をそのまま渡すことも可能です。
        output_path:  出力先ファイルパス（デフォルト: ./output/presentation.pptx）
        theme:        デザインテーマ — "aws"（デフォルト）| "minimal" | "dark"
        model:        使用する Claude モデル（デフォルト: claude-sonnet-4-6）

    Returns:
        生成されたPPTXファイルの絶対パスと、スライド枚数などのサマリー。
    """
    try:
        spec = parse_input(description, model=model)
        # Override theme if explicitly specified
        if theme and theme != "aws":
            spec = spec.model_copy(update={"theme": theme})
        path = generate(spec, output_path)
        slide_summary = "\n".join(
            f"  {i+1}. [{s.type}] {getattr(s, 'title', '')}"
            for i, s in enumerate(spec.slides)
        )
        return (
            f"✅ PPTX generated: {path.resolve()}\n"
            f"   Theme: {spec.theme}  |  Slides: {len(spec.slides)}\n\n"
            f"Slide list:\n{slide_summary}"
        )
    except Exception as exc:
        return f"❌ Error: {exc}"


# ---------------------------------------------------------------------------
# Tool: preview_spec
# ---------------------------------------------------------------------------

@mcp.tool()
def preview_spec(
    description: str,
    model: str = "claude-sonnet-4-6",
) -> str:
    """
    プレゼンテーション仕様をYAML形式でプレビューします（ファイルは生成しません）。

    スライド構成を確認・調整してから `create_presentation` に渡すワークフローで使えます。

    Args:
        description: 自然言語での説明
        model:       使用する Claude モデル（デフォルト: claude-sonnet-4-6）

    Returns:
        生成されたYAMLスペック文字列
    """
    try:
        spec = parse_input(description, model=model)
        return yaml.dump(
            spec.model_dump(exclude_none=True),
            allow_unicode=True,
            default_flow_style=False,
            sort_keys=False,
        )
    except Exception as exc:
        return f"❌ Error: {exc}"


# ---------------------------------------------------------------------------
# Tool: create_from_yaml
# ---------------------------------------------------------------------------

@mcp.tool()
def create_from_yaml(
    yaml_spec: str,
    output_path: str = "./output/presentation.pptx",
) -> str:
    """
    YAML仕様書から直接PPTXを生成します（Claude APIを使いません）。

    `preview_spec` で確認・編集したYAMLをそのまま渡してください。

    Args:
        yaml_spec:   YAML形式のスライド仕様書（文字列）
        output_path: 出力先ファイルパス

    Returns:
        生成されたPPTXファイルの絶対パス
    """
    try:
        spec = spec_from_yaml(yaml_spec)
        path = generate(spec, output_path)
        return f"✅ PPTX generated: {path.resolve()}  ({len(spec.slides)} slides)"
    except Exception as exc:
        return f"❌ Error: {exc}"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
