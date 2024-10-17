import os
from typing import Optional


def convert_excel_to_pdf(excel_file_path: str, output_file_path: Optional[str] = None) -> Optional[str]:
    """
    Функция создает из исходного excel-файла по адресу excel_file_path новый pdf-файл.
    Сохраняет новый pdf-файл по адресу output_file_path. Если output_file_path не передан,
    новый файл сохраняется в директории с исходным файлом.
    :return: Путь к созданному pdf-файлу
    """
    pass


if __name__ == '__main__':
    excel_file_path = '../../output/1_Заглушка_эллип____CAP_12__SCH_30_BD_ASTM_A420_GR_WPL6.xlsx'
    convert_excel_to_pdf(excel_file_path=excel_file_path)
