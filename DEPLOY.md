# Развёртывание EPGURU на сервере

Нужен любой Linux-сервер с Docker. Сервер с VPN не подходит: сервисы и VPN держим раздельно.

## Рекомендуемый вариант: Cloudflare Tunnel + Access (сайт закрыт для чужих)

Порты на сервере наружу не открываются: сервер сам подключается к Cloudflare. До сайта доходят только те, кто прошёл вход Cloudflare (код на почту из вашего списка). Пароль EPGURU остаётся вторым замком.

1. Cloudflare → **Zero Trust → Networks → Tunnels → Create tunnel** (тип Cloudflared). Скопируйте токен.
2. В том же туннеле **Public Hostname**: `epg.<ваш домен>` → Service `http://epguru:8000`.
3. **Zero Trust → Access → Applications → Add → Self-hosted**: домен `epg.<ваш домен>`, политика Allow → Emails: ваши адреса.
4. На сервере:
```bash
curl -fsSL https://get.docker.com | sh
git clone https://github.com/AlexCFM/EPGURU.git /opt/epguru && cd /opt/epguru
cp .env.example .env
sed -i "s/^EPGURU_SECRET=.*/EPGURU_SECRET=$(openssl rand -hex 32)/" .env
nano .env    # EPGURU_PASSWORD и CF_TUNNEL_TOKEN=<токен из шага 1>
mkdir -p data && chown -R 1000:1000 data   # приложение работает не от root
docker compose -f docker-compose.tunnel.yml up -d --build
```
Обновление: `git pull && docker compose -f docker-compose.tunnel.yml up -d --build`.

---

## Вариант с открытым сайтом (Caddy, свой HTTPS)

## 1. Домен
Для домена (например `epg.example.com`) заведите A-запись на IP сервера.
Если пока без домена, см. «Без домена» ниже.

## 2. Установка (один раз)
```bash
# Docker, если его ещё нет
curl -fsSL https://get.docker.com | sh

git clone https://github.com/AlexCFM/EPGURU.git /opt/epguru
cd /opt/epguru
cp .env.example .env
sed -i "s/^EPGURU_SECRET=.*/EPGURU_SECRET=$(openssl rand -hex 32)/" .env
nano .env          # впишите EPGURU_DOMAIN и EPGURU_PASSWORD
mkdir -p data && chown -R 1000:1000 data
docker compose up -d --build
```
Через минуту сайт откроется на `https://<домен>`. HTTPS-сертификат Caddy получает сам.

## Обновление (после правок на GitHub)
```bash
cd /opt/epguru && git pull && docker compose up -d --build
```

## Без домена (временно, по IP)
В `.env`:
```
EPGURU_DOMAIN=:80
EPGURU_SECURE_COOKIE=0
```
Сайт будет на `http://<IP>`. Соединение не шифруется, поэтому это только для проверки.

## Если порты 80/443 уже заняты (стоит nginx или другой Caddy)
Уберите сервис `caddy` из `docker-compose.yml` и пробросьте порт приложения:
```yaml
    ports:
      - "127.0.0.1:8010:8000"
```
Затем в своём прокси направьте домен на `127.0.0.1:8010` и поставьте лимит тела запроса не меньше 100 МБ.

## Данные и резервная копия
Всё хранится в `/opt/epguru/data`: база истории и файлы (исходники и результаты).
Для резервной копии достаточно скопировать эту папку.

Чистка заданий старше 90 дней (можно повесить в cron раз в сутки):
```bash
docker compose exec epguru python -c "from server.main import cleanup; print(cleanup(90))"
```
