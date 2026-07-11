import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app_info
import build


class AppInfoTests(unittest.TestCase):
    def test_version_title_and_changelog_use_current_version(self):
        self.assertEqual(
            app_info.format_version_title(),
            f"{app_info.APP_NAME} v{app_info.APP_VERSION}",
        )

        changelog = app_info.format_changelog()
        self.assertIn(f"当前版本：v{app_info.APP_VERSION}", changelog)
        self.assertIn(f"发布日期：{app_info.APP_RELEASE_DATE}", changelog)
        self.assertIn(app_info.APP_CHANGELOG[0]["items"][0], changelog)

    def test_release_info_is_valid(self):
        build._validate_release_info()
        parts = [int(part) for part in app_info.APP_VERSION.split(".")]
        expected = tuple((parts + [0, 0, 0, 0])[:4])
        self.assertEqual(build._version_tuple(), expected)

    def test_windows_version_file_uses_current_version(self):
        with TemporaryDirectory() as tmp_dir:
            version_file = Path(tmp_dir) / "version_info.txt"
            build._write_version_info_file(str(version_file))
            content = version_file.read_text(encoding="utf-8")

        self.assertIn(f"filevers={build._version_tuple()}", content)
        self.assertIn(
            f"StringStruct('ProductVersion', '{app_info.APP_VERSION}')",
            content,
        )


if __name__ == "__main__":
    unittest.main()
