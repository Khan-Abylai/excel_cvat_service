from PIL import Image, ImageDraw, ImageOps
import os
import xml.etree.ElementTree as ET
import openpyxl
import shutil

import src.config as config
import sys


def load_brands_from_excel(brands_file_path):
    try:
        workbook = openpyxl.load_workbook(brands_file_path)
    except:
        print('Unable to open brands list')
        return None
    sheet = workbook.active
    brands = [cell.value.lower() for cell in sheet['A'] if cell.value is not None]
    return brands


def parse_annotations(annotation_file_path):
    tree = ET.parse(annotation_file_path)
    root = tree.getroot()

    images_data = []
    for child in root:
        if child.tag == 'image':
            box = [subchild.attrib for subchild in child if subchild.tag == 'box']
            point = [subchild.attrib for subchild in child if subchild.tag == 'points']
            element_info = {
                'tag': child.tag,
                'attributes': child.attrib,
                'children': {'box': box, 'points': point}
            }
            images_data.append(element_info)

    return images_data


def get_unique_filename(file_path):
    base, extension = os.path.splitext(file_path)
    counter = 1
    while os.path.exists(file_path):
        file_path = f"{base}_{counter}{extension}"
        counter += 1
    return file_path


def save_crop(img_name, bbox_orig, destination):
    image_path = img_name
    original_image = Image.open(image_path)
    original_image = ImageOps.exif_transpose(original_image)  # Применяем ориентацию EXIF
    cropped_image = original_image.crop(bbox_orig)
    cropped_image.save(destination)


def visualize_annotations(image_path, annotations, output_folder, output_filename):
    """
    Визуализирует bounding boxes и точки на изображении.

    :param image_path: Путь к исходному изображению.
    :param annotations: Словарь аннотаций с ключами 'box' и 'points'.
    :param output_folder: Путь к папке для сохранения визуализированного изображения.
    :param output_filename: Имя файла для сохранения визуализированного изображения.
    """
    try:
        # Убедимся, что папка для сохранения существует
        os.makedirs(output_folder, exist_ok=True)

        output_path = os.path.join(output_folder, output_filename)

        image = Image.open(image_path).convert("RGB")
        image = ImageOps.exif_transpose(image)  # Применяем ориентацию EXIF
        draw = ImageDraw.Draw(image)

        # Рисуем bounding boxes
        for box in annotations.get('box', []):
            try:
                x1 = int(float(box['xtl']))
                y1 = int(float(box['ytl']))
                x2 = int(float(box['xbr']))
                y2 = int(float(box['ybr']))
                label = box.get('label', 'unknown')
                # Рисуем прямоугольник красным цветом
                draw.rectangle([x1, y1, x2, y2], outline="red", width=2)
                # Добавляем текст метки
                draw.text((x1, y1 - 10), label, fill="red")
            except KeyError:
                continue

        # Рисуем точки
        for point in annotations.get('points', []):
            try:
                points = list(map(int, map(float, point['points'].split(','))))
                label = point.get('label', 'unknown')
                x, y = points
                # Рисуем синий кружок
                draw.ellipse([x - 5, y - 5, x + 5, y + 5], fill="blue", outline="blue")
                # Добавляем текст метки
                draw.text((x + 10, y - 10), label, fill="blue")
            except (KeyError, ValueError):
                continue

        image.save(output_path)
        print(f"Визуализированное изображение сохранено как {output_path}")
    except Exception as e:
        print(f"Ошибка при визуализации аннотаций для {image_path}: {e}")


def process_with_filter(images_data, brands, data_folder):
    for image_data in images_data:
        checked = 0
        for point_data in image_data['children']['points']:
            point_data_coord = list(map(int, map(float, point_data['points'].split(','))))
            for box_data in image_data['children']['box']:
                box_bbox = [
                    int(float(box_data['xtl'])),
                    int(float(box_data['ytl'])),
                    int(float(box_data['xbr'])),
                    int(float(box_data['ybr']))
                ]
                if (brands is None or (point_data['label'].lower() in brands) and
                        box_bbox[0] < point_data_coord[0] < box_bbox[2] and
                        box_bbox[1] < point_data_coord[1] < box_bbox[3]):
                    box_data['label'] = point_data['label']
                    box_data['x'] = point_data_coord[0]
                    box_data['y'] = point_data_coord[1]
                    destination = get_unique_filename(
                        f'{data_folder}/cropped/cropped_{image_data["attributes"]["name"]}'
                    )
                    image_full_path = os.path.join(data_folder, 'images', image_data['attributes']['name'])
                    save_crop(image_full_path, box_bbox, destination)
                    box_data['cropped_name'] = os.path.basename(destination)
                    checked = 1
        image_data['checked'] = checked

    images_changed_data = [x for x in images_data if x['checked'] == 1]

    for image_data in images_changed_data:
        image_data['children']['box'] = [x for x in image_data['children']['box'] if x['label'].lower() != 'ignore']
        # Закомментируйте или удалите строки, которые удаляют точки
        # try:
        #     del image_data['children']['points']
        # except:
        #     pass

    return images_changed_data


def process_without_label_filter(images_data, data_folder, process_by_brands):
    for image_data in images_data:
        checked = 0
        if process_by_brands is True:
            for box_data in image_data['children']['box']:
                box_bbox = [
                    int(float(box_data['xtl'])),
                    int(float(box_data['ytl'])),
                    int(float(box_data['xbr'])),
                    int(float(box_data['ybr']))
                ]
                destination = get_unique_filename(
                    f'{data_folder}/cropped/cropped_{image_data["attributes"]["name"]}'
                )
                image_full_path = os.path.join(data_folder, 'images', image_data['attributes']['name'])
                save_crop(image_full_path, box_bbox, destination)
                box_data['cropped_name'] = os.path.basename(destination)
                checked = 1
        else:
            for point_data in image_data['children']['points']:
                point_data_coord = list(map(int, map(float, point_data['points'].split(','))))
                for box_data in image_data['children']['box']:
                    box_bbox = [
                        int(float(box_data['xtl'])),
                        int(float(box_data['ytl'])),
                        int(float(box_data['xbr'])),
                        int(float(box_data['ybr']))
                    ]
                    if (box_bbox[0] <= point_data_coord[0] <= box_bbox[2] and
                            box_bbox[1] <= point_data_coord[1] <= box_bbox[3]):
                        box_data['label'] = point_data['label']
                        box_data['x'] = point_data_coord[0]
                        box_data['y'] = point_data_coord[1]
                        destination = get_unique_filename(
                            f'{data_folder}/cropped/cropped_{image_data["attributes"]["name"]}'
                        )
                        image_full_path = os.path.join(data_folder, 'images', image_data['attributes']['name'])
                        save_crop(image_full_path, box_bbox, destination)
                        box_data['cropped_name'] = os.path.basename(destination)
                        checked = 1
        image_data['checked'] = checked

    return images_data


def process_images(images_data, brands_path, data_folder, process_by_brands):
    cropped_folder_path = os.path.join(data_folder, 'cropped')
    if os.path.exists(cropped_folder_path):
        shutil.rmtree(cropped_folder_path)
    os.makedirs(cropped_folder_path)

    brands = load_brands_from_excel(brands_path)
    cover_all = False
    if brands is None:
        images_changed_data = process_without_label_filter(images_data, data_folder, process_by_brands)
        cover_all = True
    else:
        images_changed_data = process_with_filter(images_data, brands, data_folder)

    return images_changed_data, cover_all


def pixels_to_width_units(pixels):
    return pixels / config.pixels_to_width_divider


def get_image_size_with_aspect_ratio(image_path, max_width):
    with Image.open(image_path) as img:
        original_width, original_height = img.size
        if original_width > max_width:
            scaling_factor = max_width / original_width
            new_width = max_width
            new_height = int(original_height * scaling_factor)
            return new_width, new_height
        else:
            return original_width, original_height
