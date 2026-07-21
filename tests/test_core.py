import importlib.util
import unittest
from pathlib import Path


SOURCE = Path(__file__).resolve().parents[1] / "採購單大量修改數量與總價.py"
SPEC = importlib.util.spec_from_file_location("purchase_order_updater", SOURCE)
APP = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(APP)


def make_document(second_part="ITEM-DEMO-B"):
    return f"""<html><body>
<table>
<tr><td><span>採購單號</span></td><td><span>PO-DEMO-001</span></td></tr>
<tr><td><span>1</span></td><td><span>ITEM-DEMO-A</span></td><td><span>PCS</span></td><td><span>10</span></td><td><span>2.50</span></td><td><span>25.00</span></td></tr>
<tr><td><span>2</span></td><td><span>{second_part}</span></td><td><span>PCS</span></td><td><span>20</span></td><td><span>2.00</span></td><td><span>40.00</span></td></tr>
<tr><td><span>採購金額合計</span></td><td><span>65.00</span></td></tr>
</table>
</body></html>"""


class ProcessPurchaseOrderTests(unittest.TestCase):
    def test_updates_exact_match_and_preserves_unlisted_item(self):
        document = make_document()
        ship = {
            "PO-DEMO-001|1": {
                "qty": 20.0,
                "part": "ITEM-DEMO-A",
                "matched": False,
                "seen": False,
                "valid": True,
            }
        }

        output, stats, _ = APP.process_po_html(document, ship, {"PO-DEMO-001"})

        self.assertIn("<span>20</span></td><td><span>2.50</span>", output)
        self.assertIn("<span>20</span></td><td><span>2.00</span>", output)
        self.assertIn("<span>90.00</span>", output)
        self.assertEqual(stats["updated"], 1)
        self.assertEqual(stats["skipped_not_listed"], 1)
        self.assertEqual(stats["totals"], 1)

    def test_part_mismatch_preserves_entire_document(self):
        document = make_document()
        ship = {
            "PO-DEMO-001|1": {
                "qty": 20.0,
                "part": "ITEM-DEMO-X",
                "matched": False,
                "seen": False,
                "valid": True,
            }
        }

        output, stats, _ = APP.process_po_html(document, ship, {"PO-DEMO-001"})

        self.assertEqual(output, document)
        self.assertEqual(stats["updated"], 0)
        self.assertEqual(stats["skipped_part_mismatch"], 1)
        self.assertEqual(stats["totals"], 0)


if __name__ == "__main__":
    unittest.main()
