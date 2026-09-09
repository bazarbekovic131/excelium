# Развёртывание Doc-V Gateway

Сервер: 192.168.30.19, пользователь radmin, порт 25353.
Старый сервис excelium (порт 25351) отключён — см. «Вывод старого
сервиса» ниже.

## Первая установка

```
ssh radmin@192.168.30.19
git clone <repo> ~/Documents/docv-gateway && cd ~/Documents/docv-gateway
python3.11 -m venv venv
venv/bin/pip install -r requirements.txt
cp .env.example .env && chmod 600 .env   # заполнить токены!
sudo cp deploy/docv-gateway.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now docv-gateway
```

Проверка с сервера Doc-V (192.168.30.29) — заодно проверяет IP-allowlist:

```
curl http://192.168.30.19:25353/health
```

С любой другой машины тот же запрос обязан вернуть 403.

## Обновление

```
cd ~/Documents/docv-gateway && git pull

sudo systemctl restart docv-gateway
```

## Вывод старого сервиса excelium (сделано)

Порядок важен: сначала гасится юнит, потом приходит обновление, которое
удаляет его код. Наоборот — gunicorn останется без файлов.

```
sudo systemctl disable --now excelium.service
cd ~/Documents/api_excel
cp utils/firmen_und_objekte.py ~/firmen_backup.py     # матрица нужна дальше
diff <(md5sum excel_templates/*.xlsx | awk '{print $1}') \
     <(md5sum templates/excel/*.xlsx | awk '{print $1}')   # пусто = копии совпали
git pull
python3 scripts/approvers_from_py.py > data/approvers.yaml
sudo systemctl restart docv-gateway
curl -s localhost:25353/health
```

Что уходит: `app.py`, `config.py`, `core/`, `models/`, `routes/`,
`utils/*.py` кроме матрицы, `excelium.service`, `scripts/parity_check.py`
(сверял старый рендер с новым, сверять больше не с чем). Код остаётся в
истории репозитория.

Что можно удалить руками, когда перестанет быть жалко: `saves/`
(выдача старого сервиса), `excel_templates/` (после проверки diff выше),
`myenv/` (его окружение с Flask и gunicorn). Порт 25351 не
переиспользуется.

## Что нужно для отдельных модулей

- /render/typst — бинарь `typst`. Под systemd PATH урезан
  (/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin), поэтому
  typst из `cargo install typst-cli` (он ложится в ~/.cargo/bin) в PATH не
  виден, хотя `which typst` в консоли его находит. Любой из вариантов:
  `sudo ln -sf /home/radmin/.cargo/bin/typst /usr/local/bin/typst` (симлинк
  переживает обновления cargo, в отличие от копии), либо `GW_TYPST_BIN` с
  абсолютным путём в .env, либо PATH в юните (там он уже прописан).
  Найденный путь виден на «Обзоре» интерфейса; без бинаря endpoint
  отвечает 503.
- /ops xlsx_to_pdf и blank_to_png (DOCX-бланки) — libreoffice
  (`apt install libreoffice --no-install-recommends`); PDF-бланки
  конвертируются без него (pymupdf в venv).
- /ops restart_unit — строка в sudoers (через `visudo`):
  `radmin ALL=(root) NOPASSWD: /usr/bin/systemctl restart docv-server.service`

## Данные, которых нет в git

`utils/firmen_und_objekte.py` (матрица подписантов) — персональные
данные: их правят на сервере, поэтому в репозитории их нет и быть не
должно. При развёртывании на новой машине файл переносят руками.

Шаблоны Excel, наоборот, лежат в git — `templates/excel/`. Каталог
`excel_templates/` рядом остался от отключённого сервиса excelium; шлюз
его не читает. Правка шаблона теперь идёт через репозиторий: изменили,
закоммитили, `git pull` на сервере.

Справочник шлюза `data/approvers.yaml` собирается из матрицы на месте:

```
python3 scripts/approvers_from_py.py > data/approvers.yaml
sudo systemctl restart docv-gateway
```

Делайте это после каждой правки матрицы, иначе шлюз будет печатать
прежних подписантов.

## Секреты

Токены генерируются `python3 -c "import secrets; print(secrets.token_urlsafe(32))"`.
GW_VERIFY_SECRET — секрет кода подлинности на печатных формах: код —
HMAC от данных документа, проверяется повторным формированием; менять
секрет не стоит, старые листы перестанут сходиться.
Значение GW_TOKEN_DOCV дублируется в Doc-V: карточка НАСТРОЙКИ, setting-поле с
токеном; действия «HTTP-запрос» подставляют его в заголовок Authorization из поля.
Ротация: поменять в .env, перезапустить юнит, поменять в НАСТРОЙКИ.

## Веб-интерфейс

Девять экранов по дизайн-системе ERP SHARK: рабочий стол, очередь
заданий, операции с конструктором, файлы, шаблоны Typst, рендер,
API-консоль, настройки, вход.

`http://192.168.30.19:25353/ui` — обзор (статусы, audit-лента), очередь
заданий, ручной запуск операций, хранилище файлов, тестовый рендер
реестров и Typst. Вход по GW_TOKEN_ADMIN (пустой токен = интерфейс
выключен).

Конструктор операций пишет ops.yaml и перечитывает реестр на лету —
перезапуск юнита не нужен. Команда ограничена: argv[0] обязан быть
{python}, {app_dir}/… либо путём в /usr/bin, /bin, /usr/local/bin,
/opt — иначе админ-токен стал бы равен доступу к оболочке. Настройки
меняют сроки и лимиты на лету (var/settings.json); токены и списки
адресов остаются в .env: смена их из веба могла бы закрыть доступ
самому администратору.

IP-доступ двумя списками: GW_ALLOWLIST — API, только сервер Doc-V;
GW_UI_ALLOWLIST — дополнительно /ui, /health и /files/* для машин
администраторов (оба принимают подсети; офисный диапазон 192.168.26.x–31.x — это два CIDR: «192.168.26.0/23» и «192.168.28.0/22»). Машина из
GW_UI_ALLOWLIST не достаёт до API даже с валидным токеном.
