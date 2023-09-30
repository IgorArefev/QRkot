from datetime import datetime
from typing import List

from aiogoogle import Aiogoogle

from app.core.config import settings
from app.models.charity_project import CharityProject


DATETIME_NOW = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
FILE_TITLE = f'Отчет от {DATETIME_NOW}'
LIST_TITLE = 'Отчет'
ROWS = 100
COLUMS = 10
SPREADSHEET_BODY = {
    'properties': {
        'title': FILE_TITLE,
        'locale': 'ru_RU'
    },
    'sheets': {
        'properties': {
            'sheetType': 'GRID',
            'sheetId': 0,
            'title': 'Лист1',
            'gridProperties': {
                'rowCount': ROWS,
                'columnCount': COLUMS
            }
        }
    }
}
TABLE_VALUES = [
    ['Отчет от', ],
    ['Топ проектов по скорости закрытия'],
    ['Название проекта', 'Время сбора', 'Описание']
]


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    service = await wrapper_services.discover('sheets', 'v4')
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=SPREADSHEET_BODY)
    )
    return response['spreadsheetId']


async def set_user_permissions(
        spreadsheet_id: str,
        wrapper_services: Aiogoogle
) -> None:
    permissions_body = {'type': 'user',
                        'role': 'writer',
                        'emailAddress': settings.email}
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheet_id,
            json=permissions_body,
            fields="id"
        )
    )


async def spreadsheets_update_value(
        spreadsheet_id: str,
        projects: List[CharityProject],
        wrapper_services: Aiogoogle
) -> None:
    service = await wrapper_services.discover('sheets', 'v4')

    for project in projects:
        new_row = [
            project.name,
            str(project.close_date - project.create_date),
            project.description
        ]
        TABLE_VALUES.append(new_row)

    update_body = {
        'majorDimension': 'ROWS',
        'values': TABLE_VALUES
    }

    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheet_id,
            range='A1:E100',
            valueInputOption='USER_ENTERED',
            json=update_body
        )
    )
