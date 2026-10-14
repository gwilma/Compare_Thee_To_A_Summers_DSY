from summers_dsy.charts import _spread


def test_end_labels_are_spread_apart_in_order():
    out = _spread([(30.0, "Daily max"), (29.5, "Tcomf"), (15.0, "Daily min")], gap=2.0)
    assert out == [(15.0, "Daily min"), (29.5, "Tcomf"), (31.5, "Daily max")]


def test_end_labels_already_apart_are_unchanged():
    assert _spread([(10.0, "a"), (20.0, "b")], gap=2.0) == [(10.0, "a"), (20.0, "b")]
