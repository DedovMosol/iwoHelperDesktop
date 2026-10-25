# -*- coding: utf-8 -*-
"""Тесты разбора ссылок README (tools/check_readme_links.py).

Сеть здесь не нужна: проверяется только то, КАКИЕ адреса проверка увидит. Именно
это и было сломано — цели ссылок-бейджей не попадали в список вовсе.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import check_readme_links  # noqa: E402


class ExtractTargetsTest(unittest.TestCase):
    def test_badge_link_target_is_checked(self):
        text = ("[![Installer](https://img.shields.io/badge/x.svg)]"
                "(https://github.com/o/r/releases/download/v1/setup.exe)")
        targets = check_readme_links.extract_targets(text)
        self.assertIn("https://github.com/o/r/releases/download/v1/setup.exe", targets)
        self.assertIn("https://img.shields.io/badge/x.svg", targets)

    def test_badge_with_local_target(self):
        text = "[![License](https://img.shields.io/l.svg)](LICENSE) [![P](p.svg)](docs/PRIVACY.md)"
        targets = check_readme_links.extract_targets(text)
        self.assertIn("LICENSE", targets)
        self.assertIn("docs/PRIVACY.md", targets)

    def test_plain_link_image_and_html(self):
        text = ('See [docs](docs/ARCHITECTURE.md), ![shot](docs/a.png) and '
                '<img src="docs/b.png" width="10"> or [top](#-features "title")')
        targets = check_readme_links.extract_targets(text)
        for expected in ("docs/ARCHITECTURE.md", "docs/a.png", "docs/b.png", "#-features"):
            self.assertIn(expected, targets)

    def test_image_alone_is_not_a_link_target(self):
        targets = check_readme_links.extract_targets("![shot](docs/a.png)")
        self.assertEqual(["docs/a.png"], targets)

    def test_real_readme_download_buttons_are_extracted(self):
        readme = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "README.md")
        with open(readme, encoding="utf-8") as handle:
            targets = check_readme_links.extract_targets(handle.read())
        downloads = [t for t in targets if "/releases/download/" in t]
        # Четыре кнопки: установщик и portable для x64 и x86.
        self.assertEqual(4, len(downloads), downloads)


if __name__ == "__main__":
    unittest.main()
