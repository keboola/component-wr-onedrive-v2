"""
Template Component main class.

"""

import logging
import sys

from keboola.component.base import ComponentBase
from keboola.component.exceptions import UserException
from pydantic import ValidationError

from configuration import RowConfig

logger = logging.getLogger(__name__)


class Component(ComponentBase):
    """
    Extends base class for general Python components. Initializes the CommonInterface
    and performs configuration validation.

    For easier debugging the data folder is picked up by default from `../data` path,
    relative to working directory.
    """

    def __init__(self):
        super().__init__()

    def run(self):
        """
        Main execution code.

        NOTE: this is scaffolding only — the real orchestration (mode dispatch,
        client wiring, token rotation) lands in later implementation tasks.
        """
        params = self._load_configuration()
        logger.info("Loaded configuration for mode: %s", params.mode)

    def _load_configuration(self) -> RowConfig:
        try:
            return RowConfig.model_validate(self.configuration.parameters)
        except ValidationError as e:
            error_messages = [f"{'.'.join(str(loc) for loc in err['loc'])}: {err['msg']}" for err in e.errors()]
            raise UserException(f"Invalid configuration: {'; '.join(error_messages)}") from e


"""
        Main entrypoint
"""
if __name__ == "__main__":
    try:
        comp = Component()
        # this triggers the run method by default and is controlled by the configuration.action parameter
        comp.execute_action()
    except UserException:
        logger.exception("Component failed with a user error")
        sys.exit(1)
    except Exception:
        logger.exception("Component failed with an unexpected error")
        sys.exit(2)
