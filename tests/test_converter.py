import pytest

from converter import (
    ConversionOptions,
    convert_ascii_to_fullwidth,
    convert_number_to_kanji,
    convert_numbers_in_text,
    insert_space_after_punctuation,
    replace_ellipsis,
    should_indent,
)


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("0", "零"),
        ("1", "一"),
        ("10", "十"),
        ("11", "十一"),
        ("22", "二十二"),
        ("100", "百"),
        ("101", "百一"),
        ("1000", "千"),
        ("9999", "九千九百九十九"),
    ],
)
def test_convert_number_to_kanji(source, expected):
    assert convert_number_to_kanji(source) == expected


@pytest.mark.parametrize("source", ["", "12345", "１２", "12a"])
def test_convert_number_to_kanji_rejects_unsupported_values(source):
    with pytest.raises(ValueError):
        convert_number_to_kanji(source)


def test_convert_numbers_leaves_large_values_unchanged():
    assert convert_numbers_in_text("22歳と12345円") == "二十二歳と12345円"


def test_convert_numbers_only_matches_ascii_digits():
    assert convert_numbers_in_text("12と１２") == "十二と１２"


def test_replace_ellipsis_only_replaces_three_ascii_periods():
    assert replace_ellipsis("えっと...そう…") == "えっと……そう…"


def test_convert_ascii_to_fullwidth_preserves_non_letters():
    assert convert_ascii_to_fullwidth("ABC xyz 123!") == "ＡＢＣ ｘｙｚ 123!"


def test_insert_space_after_punctuation_is_idempotent():
    assert insert_space_after_punctuation("本当！？　はい") == "本当！　？　はい"
    assert insert_space_after_punctuation("本当！　はい") == "本当！　はい"
    assert insert_space_after_punctuation("本当？』") == "本当？』"


@pytest.mark.parametrize(
    ("style_name", "text", "expected"),
    [
        ("Normal", "本文", True),
        ("Normal", "　本文", False),
        ("Normal", "「会話」", False),
        ("Normal", "（注）", False),
        ("Normal", "『書名』", False),
        ("Heading 1", "見出し", False),
        ("見出し 1", "見出し", False),
        ("Normal", "", False),
        ("Normal", "   ", False),
    ],
)
def test_should_indent(style_name, text, expected):
    assert should_indent(style_name, text) is expected


def test_options_can_disable_rules():
    options = ConversionOptions.with_disabled(["numbers", "font"])
    assert not options.enabled("numbers")
    assert not options.enabled("font")
    assert options.enabled("indent")
