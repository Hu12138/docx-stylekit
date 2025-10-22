import copy

def merge_enterprise_with_observed(enterprise_yaml: dict, observed_yaml: dict) -> dict:
    """
    简化版策略：
    - enterprise 的结构为基准；
    - observed 中对于 theme/fonts/colors/styles/numbering/page_setup：
      - 若 enterprise 存在 default，则保留 default，将 observed 值加入 options（或 mapping）；
      - 若 enterprise 缺失该字段的 options，则创建；
    - 不覆盖 default，除非后续提供“模板优先”开关。
    """
    merged = copy.deepcopy(enterprise_yaml)

    # 1) theme colors/fonts → 追加 options，不动 default
    _merge_theme(merged, observed_yaml.get("theme", {}))

    # 2) styles → 将未知样式追加到 styles.* 下（不改变默认样式），供后续手选
    _merge_styles(merged, observed_yaml.get("styles", {}))

    # 3) numbering → 记录 observed 的样式绑定，若企业未定义则加入可选 preset
    _merge_numbering(merged, observed_yaml.get("numbering", {}))

    # 4) page_setup → 对齐页边距/纸张为 options_range 内的观测点（不改默认）
    _merge_page_setup(merged, observed_yaml.get("page_setup", {}))

    # 5) headers_footers → 标记页码域存在性（给校验用）
    merged.setdefault("headers_footers_observed", observed_yaml.get("headers_footers", {}))

    return merged

def _merge_theme(merged, theme_obs):
    # colors
    colors = theme_obs.get("colors", {})
    merged.setdefault("theme", {}).setdefault("colors", {})
    for k, v in colors.items():
        merged["theme"]["colors"].setdefault(k, v)
    # fonts
    fonts = theme_obs.get("fonts", {})
    mfonts = merged["theme"].setdefault("fonts", {"major": {}, "minor": {}})
    for group in ("major", "minor"):
        for key, val in fonts.get(group, {}).items():
            mfonts[group].setdefault(key, val)

def _merge_styles(merged, styles_obs):
    # 仅把新样式挂到 merged["styles_observed"]，避免污染基线
    merged["styles_observed"] = styles_obs

def _merge_numbering(merged, numbering_obs):
    merged["numbering_observed"] = numbering_obs

def _merge_page_setup(merged, page_setup_obs):
    merged["page_setup_observed"] = page_setup_obs
