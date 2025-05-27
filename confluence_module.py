from utils import init_conf_conn

# Параметры Confluence
confluence = init_conf_conn()

# Теперь можно работать с confluence, например, получить страницу
page_id = '196784269'
page = confluence.get_page_by_id(page_id)
print(page['title'])
