import pytest

from scripts import ppt_paste


class FakeShape:
    def __init__(self, name, type_):
        self.Name, self.Type = name, type_


class FakePasted:
    def __init__(self, first):
        self.first, self.deleted = first, False

    def __call__(self, i):
        return self.first

    def Delete(self):
        self.deleted = True


class FakeRange:
    def Copy(self):
        pass


class FakeArchiveShapes:
    def Range(self):
        return FakeRange()


class FakeSlideShapes:
    def __init__(self, results):
        self.results, self.pasted = list(results), []

    def Paste(self):
        p = FakePasted(self.results.pop(0))
        self.pasted.append(p)
        return p


class FakeSlide:
    def __init__(self, results):
        self.Shapes = FakeSlideShapes(results)


@pytest.fixture(autouse=True)
def no_sleep(monkeypatch):
    monkeypatch.setattr(ppt_paste.time, "sleep", lambda _: None)


def test_retries_when_clipboard_holds_previous_song():
    slide = FakeSlide([FakeShape("Старая песня", 13), FakeShape("Нужная песня", 16)])

    pasted = ppt_paste._copy_paste_checked(FakeArchiveShapes(), slide, "Нужная песня", 16)

    assert pasted(1).Name == "Нужная песня"
    assert slide.Shapes.pasted[0].deleted  # неверная вставка убрана со слайда
    assert not slide.Shapes.pasted[1].deleted


def test_gives_up_after_three_wrong_pastes():
    slide = FakeSlide([FakeShape("Старая песня", 13)] * 3)

    with pytest.raises(RuntimeError, match="Нужная песня"):
        ppt_paste._copy_paste_checked(FakeArchiveShapes(), slide, "Нужная песня", 16)

    assert all(p.deleted for p in slide.Shapes.pasted)
