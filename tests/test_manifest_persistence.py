import json
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

import input_files


class ManifestPersistenceTests(unittest.TestCase):
    def test_independent_concurrent_updates_do_not_lose_each_other(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            manifest_path = Path(temp_dir) / "manifest.json"
            manifest_path.write_text('{"files": {}, "marker": 1}', encoding="utf-8")
            rule = {
                "enabled": True,
                "priority": ["Plus", "Flaş", "Avantajlı"],
            }
            selections = {"123": "Plus"}
            first_writer_entered = threading.Event()
            release_first_writer = threading.Event()
            first_write_finished = threading.Event()
            second_writer_entered = threading.Event()
            writer_count_lock = threading.Lock()
            writer_count = 0
            errors = []
            original_write = input_files._write_manifest_atomic

            def coordinated_write(path, manifest):
                nonlocal writer_count
                with writer_count_lock:
                    writer_count += 1
                    position = writer_count
                if position == 1:
                    first_writer_entered.set()
                    if not release_first_writer.wait(2):
                        raise AssertionError("İlk manifest yazımı serbest bırakılmadı.")
                    try:
                        original_write(path, manifest)
                    finally:
                        first_write_finished.set()
                    return

                second_writer_entered.set()
                if not first_write_finished.wait(2):
                    raise AssertionError("İkinci manifest yazımı ilkini beklemedi.")
                original_write(path, manifest)

            def run(target, *args):
                try:
                    target(*args)
                except BaseException as exc:  # Thread failures must fail the test.
                    errors.append(exc)

            with patch.object(input_files, "_write_manifest_atomic", coordinated_write):
                rule_thread = threading.Thread(
                    target=run,
                    args=(input_files.save_recommendation_rule, manifest_path, rule),
                )
                rule_thread.start()
                self.assertTrue(first_writer_entered.wait(2))

                selection_thread = threading.Thread(
                    target=run,
                    args=(input_files.save_user_selections, manifest_path, selections),
                )
                selection_thread.start()
                second_writer_entered.wait(0.5)
                release_first_writer.set()

                rule_thread.join(3)
                selection_thread.join(3)

            self.assertFalse(rule_thread.is_alive())
            self.assertFalse(selection_thread.is_alive())
            self.assertEqual(errors, [])
            saved = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(saved["recommendation_rule"], rule)
            self.assertEqual(saved["user_selections"], selections)
            self.assertEqual(saved["marker"], 1)


if __name__ == "__main__":
    unittest.main()
