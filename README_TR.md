# Fortress Delta Web Dashboard v1

Bu web uygulaması Bloomberg PC'deki collector'dan yalnızca türetilmiş risk metriklerini alır.

## Ortam değişkenleri
- INGEST_TOKEN: Collector'ın /api/ingest endpoint'ine bağlanırken kullanacağı uzun gizli token.
- VIEW_USER: Portal görüntüleme kullanıcı adı.
- VIEW_PASSWORD: Portal görüntüleme parolası.
- STALE_AFTER_SECONDS: Collector verisi kaç saniye gelmezse OFFLINE gösterileceği. Varsayılan 20.

## Endpointler
- POST /api/ingest — Collector için Bearer token korumalı.
- GET / — Basic Auth korumalı dashboard.
- GET /api/snapshot — Basic Auth korumalı JSON.
- GET /health — Sağlık kontrolü.

## Render üzerinde hızlı deployment
1. Bu klasörü private bir GitHub repository'ye yükleyin.
2. Render > New > Web Service > GitHub repository'yi seçin.
3. Runtime: Python
4. Build command: pip install -r requirements.txt
5. Start command: uvicorn app:app --host 0.0.0.0 --port $PORT
6. Environment variables ekleyin:
   INGEST_TOKEN=<uzun rastgele token>
   VIEW_USER=fuat
   VIEW_PASSWORD=<güçlü portal parolası>
   STALE_AFTER_SECONDS=20
7. Deploy edin.
8. Render size bir HTTPS URL verir. Örnek:
   https://fortress-delta-dashboard.onrender.com
9. Collector içindeki Portal ingest URL alanına:
   https://fortress-delta-dashboard.onrender.com/api/ingest
   yazın.
10. Collector'a aynı INGEST_TOKEN değerini girin.

## Token üretme
Bilgisayarınızda:
python generate_secret.py

çıktısını kopyalayın. Bu token Bloomberg parolası değildir.

## Not
Bu prototip yalnızca son snapshot'ı RAM'de tutar. Server yeniden başlarsa sayfa collector'ın bir sonraki push'una kadar WAITING gösterir. Bu canlı monitor için bilinçli bir tasarımdır.
