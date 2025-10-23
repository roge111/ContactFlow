# Пример local_settings
# Измените данные на свои

DEBUG = True
ALLOWED_HOSTS = ['*']

from integration_utils.bitrix24.local_settings_class import LocalSettingsClass

TINKOFF_API_KEY = 'your-api-key'
ENDPOINT_TINKOFF = 'your-secret-key'
API_KEY_TINKOFF = 'your-api-key'
SECRET_KEY_TINKOFF = 'your-secret-key'

OPEN_AI_API_KEY = 'your-api-key'
NGROK_URL = 'http://localhost:8000/'

APP_SETTINGS = LocalSettingsClass(
    portal_domain='b24-333x61.bitrix24.ru',
    app_domain='127.0.0.1:8000',
    app_name='application-1',
    salt='wefiewofioiI(IF(Eufrew8fju8ewfjhwkefjlewfjlJFKjewubhybfwybgybHBGYBGF',
    secret_key='wefewfkji4834gudrj.kjh237tgofhfjekewf.kjewkfjeiwfjeiwjfijewf',
    application_bitrix_client_id='local.68e8d7ad225af7.22855713',
    application_bitrix_client_secret='V5tEfSlKvH08ztj37yqgeTbDcYSNpG8GVjA1QWt6Kpxl6WLpOE',
    application_index_path='/',
)

DOMAIN = "56218ef983f3-8301993767665431593.ngrok-free.app"


DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql_psycopg2',
        'NAME': 'is_demo',  # Or path to database file if using sqlite3.
        'USER': 'postgres',  # Not used with sqlite3.
        'PASSWORD': 'Gb%v5oVA',  # Not used with sqlite3.
        'HOST': 'localhost',
    },
}
