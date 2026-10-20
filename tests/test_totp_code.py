from nmbrs_uren_invullen import totp_code

# Testvectoren uit RFC 6238 (SHA1, sleutel "12345678901234567890"), laatste 6 cijfers
SECRET = "GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ"


def test_totp_rfc_vectoren():
    assert totp_code(SECRET, 59) == "287082"
    assert totp_code(SECRET, 1111111109) == "081804"
    assert totp_code(SECRET, 2000000000) == "279037"


def test_totp_spaties_en_kleine_letters():
    assert totp_code("gezd gnbv gy3t qojq gezd gnbv gy3t qojq", 59) == "287082"


def test_totp_zonder_padding():
    # Authenticator-sleutels hebben vaak een lengte die geen veelvoud van 8 is
    assert len(totp_code("JBSWY3DPEHPK3PX", 59)) == 6
