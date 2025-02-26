import logging

# Создаем логгер
logger = logging.getLogger("PresentationAppLogger")
logger.setLevel(logging.INFO)  # Устанавливаем минимальный уровень логирования

# Создаем обработчик для вывода в консоль
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# Устанавливаем формат сообщений
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
console_handler.setFormatter(formatter)

# Добавляем обработчик к логгеру
logger.addHandler(console_handler)