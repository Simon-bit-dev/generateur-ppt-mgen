class MgenGeneratorError(Exception):
    """Base exception for generator failures."""


class ValidationError(MgenGeneratorError):
    """Raised when input YAML or template violates the contract."""


class GenerationError(MgenGeneratorError):
    """Raised when the PowerPoint cannot be generated safely."""
