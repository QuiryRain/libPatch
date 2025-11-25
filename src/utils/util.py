#!/usr/bin/env python3
# -*- coding: utf8 -*-
import pandas as pd
from io import BytesIO

from lib.xlsxwriterlib import WorkbookLib


def generate_excel_binary_data(filename, data):
    workbook = WorkbookLib(filename, options={'strings_to_urls': False, 'in_memory': True})
    for key, values in data.items():
        worksheet = workbook.add_worksheet(key)
        bold = workbook.add_format({'bold': True})
        for index1, value1 in enumerate(values):
            for index2, value2 in enumerate(value1):
                if not isinstance(value2, tuple) and pd.isnull(value2):
                    value2 = ""
                if index1 == 0:
                    worksheet.write(index1, index2, value2, bold)
                elif isinstance(value2, tuple):
                    try:
                        image_name, content_buffer = value2
                        worksheet.wps_embed_image(
                            index1, index2, image_name, {"image_data": BytesIO(content_buffer)}
                        )
                    except Exception as e:
                        worksheet.write(index1, index2, None)
                else:
                    worksheet.write(index1, index2, value2)
    workbook.close()
    return