import secrets
print("APP_SECRET="+secrets.token_urlsafe(48))
print("COLLECTOR_TOKEN="+secrets.token_urlsafe(48))
