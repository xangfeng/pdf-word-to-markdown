"""视觉公式识别引擎（辅助增强，默认关闭）

支持多家视觉模型厂商，填 API key 即可启用：

  OpenAI:      模型 gpt-4o-mini / gpt-4o
  DeepSeek:    模型 deepseek-v4-pro / deepseek-v4-flash (需视觉权限)
  Kimi:        模型 moonshot-v1-8k-vision-preview / kimi-k2.6
  自定义:      任意 OpenAI 兼容接口

用法:
  from pdf_word_to_markdown.engines import FormulaVisionEngine
  fe = FormulaVisionEngine(provider="kimi", api_key="sk-xxx")
  fe.enhance_markdown(md_text, images_dir)

配置 .env:
  MOONSHOT_API_KEY=sk-xxx    # Kimi
  OPENAI_API_KEY=sk-xxx      # OpenAI
  DEEPSEEK_API_KEY=sk-xxx    # DeepSeek
  ANTHROPIC_API_KEY=sk-xxx   # Claude
"""

import base64
import os
import re
from typing import Optional


# ─── 厂商预设 ────────────────────────────────────────────────────────────

PROVIDERS = {
    "openai": {
        "name": "OpenAI",
        "base_url": "https://api.openai.com/v1",
        "models": ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini"],
        "default_model": "gpt-4o-mini",
        "env_key": "OPENAI_API_KEY",
        "setup_note": "获取 key: https://platform.openai.com/api-keys",
    },
    "deepseek": {
        "name": "DeepSeek",
        "base_url": "https://api.deepseek.com",
        "models": ["deepseek-v4-pro", "deepseek-v4-flash"],
        "default_model": "deepseek-v4-flash",
        "env_key": "DEEPSEEK_API_KEY",
        "setup_note": "获取 key: https://platform.deepseek.com/api_keys (需开通视觉权限)",
    },
    "kimi": {
        "name": "Kimi (Moonshot)",
        "base_url": "https://api.moonshot.cn/v1",
        "models": ["moonshot-v1-8k-vision-preview", "moonshot-v1-32k-vision-preview", "kimi-k2.6"],
        "default_model": "moonshot-v1-8k-vision-preview",
        "env_key": "MOONSHOT_API_KEY",
        "setup_note": "获取 key: https://platform.kimi.com",
    },
    "anthropic": {
        "name": "Anthropic Claude",
        "base_url": "https://api.anthropic.com",
        "models": ["claude-haiku-4-5", "claude-sonnet-4-6", "claude-opus-4-7"],
        "default_model": "claude-haiku-4-5",
        "env_key": "ANTHROPIC_API_KEY",
        "setup_note": "获取 key: https://console.anthropic.com (使用 Anthropic SDK 而非 OpenAI)",
    },
    "custom": {
        "name": "自定义 (OpenAI 兼容)",
        "base_url": "",
        "models": [],
        "default_model": "",
        "env_key": "CUSTOM_VISION_API_KEY",
        "setup_note": "设置 CUSTOM_VISION_API_KEY, CUSTOM_VISION_BASE_URL, CUSTOM_VISION_MODEL",
    },
}


class FormulaVisionEngine:
    """多厂商视觉公式识别 —— 公式图片 → LaTeX ($...$ / $$...$$)

    用法:
        fe = FormulaVisionEngine(provider="kimi")
        # 或自定义:
        fe = FormulaVisionEngine(provider="custom", base_url="https://api.xxx.com/v1", model="xxx-vision")
    """

    def __init__(
        self,
        provider: str = "kimi",
        api_key: str | None = None,
        base_url: str | None = None,
        model: str | None = None,
    ):
        cfg = PROVIDERS.get(provider, PROVIDERS["custom"])
        self.provider = provider

        # API key: 参数 > 环境变量 > No
        env_key = cfg.get("env_key", "")
        self.api_key = api_key or os.environ.get(env_key, "")
        self.base_url = base_url or cfg.get("base_url", "")
        self.model = model or cfg.get("default_model", "")

        if provider == "custom" and not self.base_url:
            self.base_url = os.environ.get("CUSTOM_VISION_BASE_URL", "")
        if provider == "custom" and not self.model:
            self.model = os.environ.get("CUSTOM_VISION_MODEL", "")

        self._client = None
        self._is_anthropic = (provider == "anthropic")

    @property
    def ready(self) -> bool:
        """是否已配置好，可以调用"""
        return bool(self.api_key and self.base_url and self.model)

    def _get_openai_client(self):
        if self._client is None:
            from openai import OpenAI
            self._client = OpenAI(api_key=self.api_key, base_url=self.base_url)
        return self._client

    def image_to_latex(self, image_path: str) -> Optional[str]:
        """单张公式图片 → LaTeX"""
        if not self.ready:
            return None

        ext = os.path.splitext(image_path)[1].lstrip(".") or "png"
        try:
            with open(image_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            return None

        prompt = (
            "将这个数学公式转为LaTeX格式。"
            "如果是独占一行的块级公式用$$...$$包裹，"
            "如果是文中的行内公式用$...$包裹。"
            "只输出LaTeX代码，不要任何解释。"
        )

        if self._is_anthropic:
            return self._call_anthropic(b64, ext, prompt)
        else:
            return self._call_openai_compatible(b64, ext, prompt)

    def _call_openai_compatible(self, b64: str, ext: str, prompt: str) -> Optional[str]:
        client = self._get_openai_client()
        try:
            resp = client.chat.completions.create(
                model=self.model,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image_url", "image_url": {"url": f"data:image/{ext};base64,{b64}"}},
                        {"type": "text", "text": prompt},
                    ],
                }],
                max_tokens=500,
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            print(f"  [formula-vision/{self.provider}] {e}")
            return None

    def _call_anthropic(self, b64: str, ext: str, prompt: str) -> Optional[str]:
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=self.api_key)
            resp = client.messages.create(
                model=self.model,
                max_tokens=500,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": f"image/{ext}", "data": b64}},
                        {"type": "text", "text": prompt},
                    ],
                }],
            )
            return resp.content[0].text.strip()
        except Exception as e:
            print(f"  [formula-vision/anthropic] {e}")
            return None

    def enhance_markdown(self, md_text: str, images_dir: str) -> str:
        """扫描 MD 中的公式图片，用视觉 API 增强为 LaTeX"""
        if not self.ready:
            return md_text

        img_pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')

        def replacer(match):
            alt = match.group(1)
            src = match.group(2)
            kw = ['formula', '公式', 'equation', 'math']
            if not any(k in alt.lower() or k in src.lower() for k in kw):
                return match.group(0)

            for candidate in [
                os.path.join(images_dir, os.path.basename(src)),
                src if os.path.exists(src) else None,
            ]:
                if candidate and os.path.exists(candidate):
                    latex = self.image_to_latex(candidate)
                    if latex:
                        return latex
                    break
            return match.group(0)

        return img_pattern.sub(replacer, md_text)


# ─── 快捷函数 ────────────────────────────────────────────────────────────


def create_engine(provider: str = "kimi", **kwargs) -> FormulaVisionEngine:
    """快速创建引擎"""
    return FormulaVisionEngine(provider=provider, **kwargs)


def list_providers() -> list[str]:
    """列出所有支持的厂商"""
    return list(PROVIDERS.keys())


def get_provider_info(provider: str) -> dict:
    """获取厂商配置信息"""
    return PROVIDERS.get(provider, PROVIDERS["custom"])
