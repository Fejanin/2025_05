import os
from PIL import Image


def combin_png(img_head, img_data, name):
    img_top_name = img_head
    img_bottom_name = img_data

    # Открываем изображения
    img_top = Image.open(img_top_name)
    img_bottom = Image.open(img_bottom_name)

    # Вертикальное объединение: высота — сумма, ширина — максимум
    total_width = max(img_top.width, img_bottom.width)
    total_height = img_top.height + img_bottom.height

    # Создаём пустое изображение для результата
    result = Image.new("RGB", (total_width, total_height), color="white")

    # Вставляем части
    result.paste(img_top, (0, 0))
    result.paste(img_bottom, (0, img_top.height))  # , (img1.width, 0)) - объединение по горизонтали

    # Сохраняем результат
    result.save(f"result\\{name}.png")


if __name__ == "__main__":
    combin_png("head.png", "test.png", "099")
