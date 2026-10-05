from msme_verifier.validators import detect_id_type, is_valid_pan, is_valid_udyam, normalise, pan_holder_type


def test_pan_validation():
    assert is_valid_pan("AAAPL1234C")
    assert is_valid_pan(" aaapl1234c ")
    assert not is_valid_pan("AAAPL1234")
    assert not is_valid_pan("1AAPL1234C")


def test_udyam_validation():
    assert is_valid_udyam("UDYAM-MH-26-0012345")
    assert is_valid_udyam("udyam mh 26 0012345".replace(" ", "-"))
    assert not is_valid_udyam("UDYAM-MH-26-001234")
    assert not is_valid_udyam("UDYAM-M1-26-0012345")


def test_detect_and_normalise():
    assert normalise(" udyam_mh_26_0012345 ") == "UDYAM-MH-26-0012345"
    assert detect_id_type("UDYAM-MH-26-0012345") == "UDYAM"
    assert detect_id_type("AAAPL1234C") == "PAN"
    assert detect_id_type("hello") is None
    assert detect_id_type(None) is None


def test_pan_holder_type():
    assert pan_holder_type("AAAPL1234C") == "Individual / Proprietorship"
    assert pan_holder_type("AADCS4321M") == "Company"
    assert pan_holder_type("bad") is None
