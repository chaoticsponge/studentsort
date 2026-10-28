import hashlib
import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from urllib.error import HTTPError

import app_updates as updates


class UpdateTests(unittest.TestCase):
    def release(self, version="99.0.0"):
        return {"tag_name": "v" + version, "assets": [{
            "name": updates.ASSET_NAME,
            "browser_download_url": f"https://github.com/{updates.REPOSITORY}/releases/download/v{version}/{updates.ASSET_NAME}",
            "digest": "sha256:" + "a" * 64, "size": 10,
        }]}

    def test_version_comparison(self):
        self.assertGreater(updates.version_tuple("v1.10.0"), updates.version_tuple("1.9.0"))
        self.assertIsNone(updates.release_update(self.release(updates.VERSION)))
        self.assertIsNotNone(updates.release_update(self.release()))

    def test_reject_untrusted_or_unverifiable_release(self):
        for field, value in [("browser_download_url", "https://example.org/app.exe"),
                             ("digest", None), ("size", updates.MAX_DOWNLOAD + 1)]:
            release = self.release()
            release["assets"][0][field] = value
            with self.assertRaises(ValueError):
                updates.release_update(release)

    def test_prerelease_and_no_release(self):
        release = self.release()
        release["prerelease"] = True
        self.assertIsNone(updates.release_update(release))
        with patch.object(updates, "open_url", side_effect=HTTPError("url", 404, "Not found", {}, None)):
            self.assertIsNone(updates.check_for_update())

    def test_download_checksum_and_cleanup(self):
        payload = b"MZsynthetic executable for checksum test"
        update = {"url": "test", "size": len(payload), "sha256": hashlib.sha256(payload).hexdigest()}
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp) / "download"
            folder.mkdir()
            with patch.object(updates.tempfile, "mkdtemp", return_value=str(folder)), \
                    patch.object(updates, "open_url", return_value=io.BytesIO(payload)):
                result = updates.download_update(update)
                self.assertEqual(result.read_bytes(), payload)
            for bad_payload in [payload[:-1], b"ZZ" + payload[2:], payload + b"extra"]:
                with patch.object(updates.tempfile, "mkdtemp", return_value=str(folder)), \
                        patch.object(updates, "open_url", return_value=io.BytesIO(bad_payload)):
                    with self.assertRaises(ValueError):
                        updates.download_update(update)
                self.assertFalse(folder.exists())
                folder.mkdir()

    def test_source_cannot_install(self):
        with patch.object(updates.sys, "frozen", False, create=True):
            with self.assertRaises(ValueError):
                updates.start_install(Path("example.exe"))


if __name__ == "__main__":
    unittest.main()
