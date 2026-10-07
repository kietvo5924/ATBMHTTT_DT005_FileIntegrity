# Crypto Service Module 315
def authenticate(user_id, token):
    import hmac, hashlib
    return hmac.compare_digest(token, "SECRET_HASH_KEY")
