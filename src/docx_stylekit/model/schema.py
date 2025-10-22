# 这里不做严格 schema 校验（先跑通），后续可引入 pydantic/cerberus
# 主要提供路径常量（给 emit/merge 用）
OBSERVED_TOP_KEYS = ["theme", "styles", "numbering", "page_setup", "headers_footers"]
