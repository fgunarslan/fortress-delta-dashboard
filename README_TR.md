# Fortress Delta Web Dashboard v2

## Yetkiler
- `fuat` sabit ve tek admin.
- Fuat viewer kullanıcı ekleyebilir, disable/enable edebilir, parola resetleyebilir ve silebilir.
- Başka admin hesabı oluşturulamaz.
- Viewer yalnızca canlı dashboard'u görür.
- Portfolio/NAV/limit/Nirvana import işlemleri yalnızca Fuat'tadır.
- Admin işlemleri Audit Log'a kaydedilir.

## Uzaktan portföy yönetimi
Collector local positions.csv kullanmaz.
Portal `/api/collector/config` endpoint'inden aktif portföy, NAV, limit ve config version yayınlar.
Her değişiklikte version artar; Bloomberg PC collector yeni versiyonu otomatik çeker.

## Nirvana import
Admin > Nirvana Import:
1. CSV/XLSX yükle.
2. Kolonları eşle.
3. Added / Changed / Missing preview'ını gör.
4. Replace veya Merge seç.
5. Confirm & Publish.

Replace: dosyada olmayan aktif pozisyonlar archive edilir.
Merge: dosyada olmayan pozisyonlara dokunulmaz.

## Expiry
Expiry tarihi bugünden eski olan pozisyon collector config'ine otomatik gönderilmez; audit/history'de kalır.

## Render Environment Variables
DATABASE_URL=<kalıcı PostgreSQL connection string>
APP_SECRET=<generate_secrets.py çıktısı>
COLLECTOR_TOKEN=<generate_secrets.py çıktısı>
ADMIN_PASSWORD=<Fuat admin parolası>
STALE_AFTER_SECONDS=20

Gerçek kullanımda kalıcı PostgreSQL kullanın. Render ephemeral filesystem üzerinde SQLite, deploy/restart sonrası veri kaybı yaratabilir.
