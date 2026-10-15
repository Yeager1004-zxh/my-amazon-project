"""领星 API 扩展占位。

当前没有已确认的领星 API 文档、权限或接口范围，因此不构造 URL、不发起请求，
也不接受或保存未经验证的密钥。获得官方授权和接口文档后再实现此适配器。
"""


class LingxingApiUnavailable(RuntimeError):
    """API 适配尚未配置。"""


def create_client() -> None:
    raise LingxingApiUnavailable(
        "领星 API 尚未配置。请先提供已授权的官方 API 文档和权限；当前请使用 Excel/CSV 导入。"
    )
