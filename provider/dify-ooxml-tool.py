import logging
from typing import Any

from dify_plugin import ToolProvider
from dify_plugin.errors.tool import ToolProviderCredentialValidationError
from dify_plugin.config.logger_format import plugin_logger_handler

# Initialize logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(plugin_logger_handler)


class DifyOoxmlToolProvider(ToolProvider):
    def _validate_credentials(self, credentials: dict[str, Any]) -> None:
        """
        OOXML Translation Tool does not require external credentials.
        This plugin processes user-uploaded files using local parsing libraries.
        """
        logger.info("[DifyOoxmlToolProvider] Validating credentials - no external credentials required")
        # No validation needed as this plugin doesn't use external APIs
        pass
