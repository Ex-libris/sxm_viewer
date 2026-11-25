import numpy as np

from sxm_viewer.utils import animation as anim


def test_playback_sequence_modes():
    assert anim.playback_sequence(4, "loop") == [0, 1, 2, 3]
    assert anim.playback_sequence(4, "once") == [0, 1, 2, 3]
    assert anim.playback_sequence(4, "pingpong") == [0, 1, 2, 3, 2, 1]
    assert anim.playback_sequence(2, "pingpong") == [0, 1]
    assert anim.playback_sequence(0, "loop") == []


def test_durations_hold_the_last_frame():
    seq = anim.playback_sequence(3, "loop")
    assert anim.frame_durations_ms(seq, 5, 1.0) == [200.0, 200.0, 1200.0]
    # ping-pong: the turnaround frame is held
    seq = anim.playback_sequence(3, "pingpong")
    assert anim.frame_durations_ms(seq, 10, 0.5) == [100.0, 100.0, 600.0, 100.0]
    assert anim.total_duration_s(3, 5, 1.0) == 1.6


def test_constant_rate_repeats_held_frame():
    seq = anim.playback_sequence(3, "loop")
    assert anim.constant_rate_sequence(seq, 4, 0.5) == [0, 1, 2, 2, 2]
    assert anim.constant_rate_sequence(seq, 4, 0.0) == [0, 1, 2]


def test_pad_to_common_centres_and_rounds():
    a = np.zeros((5, 7, 3), np.uint8)
    b = np.full((6, 4), 9, np.uint8)  # grey 2-D frame -> RGB
    out = anim.pad_to_common([a, b], fill=(255, 0, 0), multiple=2)
    assert all(f.shape == (6, 8, 3) for f in out)
    assert out[0][5, 7].tolist() == [255, 0, 0]        # padding colour (odd remainder at the far side)
    assert out[0][0, 0].tolist() == [0, 0, 0]          # original pixels kept
    assert out[1][0, 1].tolist() == [255, 0, 0]        # 4 px wide frame centred in 8
    assert out[1][3, 3].tolist() == [9, 9, 9]


def test_shared_range():
    assert anim.shared_range([(0, 1), (-1, 0.5), (None, 2)]) == (-1.0, 1.0)
    assert anim.shared_range([(1, 1)]) is None
    assert anim.shared_range([]) is None
