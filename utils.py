from loguru import logger

def setup_logging():
    """Configure centralized logging with loguru for the converter script."""
    logger.remove()  # Удаляем стандартный обработчик
    # Логи в файл
    logger.add(
        sink="logs/converter.log",
        level="DEBUG",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {name} | {message}",
        rotation="10 MB",
    )
    # Логи в консоль
    logger.add(
        sink=lambda msg: print(msg, end=""),
        level="DEBUG",
        colorize=True,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level}</level> | <cyan>{name}</cyan> | <level>{message}</level>"
    )
    return logger