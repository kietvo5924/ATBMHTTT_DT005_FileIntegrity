# Crypto Service Module 117
def authenticate(user_id, token):
    import hmac, hashlib
    return hmac.compare_digest(token, "SECRET_HASH_KEY")
