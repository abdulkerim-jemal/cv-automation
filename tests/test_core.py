import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).parents[1]))

from PIL import Image
from cv_core.image_processing import prepare_cv_image
from ocr.pipeline import normalize_date_token, mrz_checksum, parse_td3


def test_date_normalization():
    assert normalize_date_token('05 JAN 1990')=='05/01/1990'
    assert normalize_date_token('1990-01-05')=='05/01/1990'
    assert normalize_date_token('31/02/2020') is None


def test_contain_keeps_full_subject_centered():
    src = Image.new("RGB", (100, 200), (10, 20, 30))
    src.putpixel((0, 0), (0, 255, 0))
    src.putpixel((99, 199), (0, 0, 255))
    out = prepare_cv_image(src, 200, 200, mode="contain")
    assert out.size == (200, 200)
    assert out.getpixel((0, 0)) == (255, 255, 255)
    assert out.getpixel((50, 0)) == (0, 255, 0)
    assert out.getpixel((149, 199)) == (0, 0, 255)


def test_mrz_checksum_and_td3_fields():
    # Build a valid TD3 second line for deterministic unit testing.
    line1='P<UTOERIKSSON<<ANNA<MARIA<<<<<<<<<<<<<<<<<<<'
    body='L898902C<3UTO6908061F9406236ZE184226B<<<<<1'
    assert len(body)==43
    # Calculate the final composite checksum for the body components.
    composite=body[0:10]+body[13:20]+body[21:28]+body[28:43]
    line2=body+str(mrz_checksum(composite))
    result=parse_td3(line1,line2)
    assert result.valid_number and result.valid_dob and result.valid_expiry and result.valid_composite
    assert result.passport_no=='L898902C'
    assert result.dob=='06/08/1969'
    assert result.expiry=='23/06/2094'


def test_mrz_score_degrades_when_checksum_is_wrong():
    r=parse_td3('P<UTOERIKSSON<<ANNA<MARIA<<<<<<<<<<<<<<<<<<<','L898902C<3UTO6908061F9406236ZE184226B<<<<<10')
    assert r.valid_score < 90
