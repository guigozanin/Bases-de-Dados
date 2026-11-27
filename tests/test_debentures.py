import datetime as dt
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd
import requests
from openpyxl import load_workbook

import debentures


class DebenturesTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.data = dt.date(2026, 11, 25)

    @patch("debentures.requests.get")
    def test_baixa_e_consolida_como_texto(self, get):
        get.return_value = Mock(
            status_code=200,
            content="título\nsubtítulo\nCódigo@Valor\n=1+1@123\nABC@456\n".encode("latin1"),
        )
        downloads = self.root / "downloads"
        arquivo = downloads / "261125.xlsx"
        base = self.root / "bases" / "base.xlsx"
        self.assertTrue(debentures.baixar_data(self.data, arquivo))
        get.assert_called_once_with(
            debentures.BASE_URL + "261125.txt", headers=debentures.HEADERS, timeout=60
        )
        debentures.consolidar(downloads, base)
        df = pd.read_excel(base)
        self.assertEqual(df["dt_referencia"].tolist(), ["2026-11-25", "2026-11-25"])
        self.assertEqual(df["Código"].tolist(), ["=1+1", "ABC"])
        for path in (arquivo, base):
            workbook = load_workbook(path)
            self.assertEqual(workbook.active["B2"].data_type, "s")
            workbook.close()

    @patch("debentures.requests.get")
    def test_indisponivel_e_erro_nao_criam_arquivo(self, get):
        arquivo = self.root / "261125.xlsx"
        get.return_value = Mock(status_code=404)
        self.assertFalse(debentures.baixar_data(self.data, arquivo))
        get.return_value = Mock(status_code=200, content=b"tela de error")
        with self.assertRaises(ValueError):
            debentures.baixar_data(self.data, arquivo)
        get.side_effect = requests.Timeout()
        with self.assertRaises(requests.Timeout):
            debentures.baixar_data(self.data, arquivo)
        self.assertFalse(arquivo.exists())

    def test_consolidacao_ordena_remove_duplicatas_e_preserva_base_em_erro(self):
        downloads = self.root / "downloads"
        base = self.root / "base.xlsx"
        debentures.salvar_excel(pd.DataFrame([
            {"dt_referencia": "2026-11-26", "codigo": "B"},
            {"dt_referencia": "2026-11-25", "codigo": "A"},
            {"dt_referencia": "2026-11-25", "codigo": "A"},
        ]), downloads / "261126.xlsx", "Debentures")
        debentures.consolidar(downloads, base)
        self.assertEqual(pd.read_excel(base)["codigo"].tolist(), ["A", "B"])
        original = base.read_bytes()
        (downloads / "corrompido.xlsx").write_bytes(b"erro")
        with self.assertRaises(Exception):
            debentures.consolidar(downloads, base)
        self.assertEqual(base.read_bytes(), original)

    def test_calendario_anbima_disponivel(self):
        calendario = debentures.Calendar.load(name="ANBIMA")
        self.assertFalse(calendario.isbizday(dt.date(2026, 11, 29)))
        self.assertIsInstance(pd.Timestamp(calendario.offset(self.data, -6)).date(), dt.date)

    def test_main_nao_baixa_novamente_arquivos_existentes(self):
        downloads = self.root / "downloads"
        base = self.root / "bases" / "base.xlsx"
        agora = dt.datetime(2026, 11, 25, 12, 0)
        calendario = Mock()
        calendario.offset.return_value = self.data
        calendario.isbizday.return_value = True
        response = Mock(status_code=200, content=b"titulo\nsubtitulo\nCodigo@Valor\nABC@123\n")
        with patch.object(debentures, "DOWNLOADS", downloads), \
             patch.object(debentures, "BASE", base), \
             patch.object(debentures, "dt", SimpleNamespace(
                 datetime=Mock(now=Mock(return_value=agora)), timedelta=dt.timedelta
             )), \
             patch("debentures.Calendar.load", return_value=calendario), \
             patch("debentures.requests.get", return_value=response) as get:
            debentures.main()
            original = base.read_bytes()
            debentures.main()
        get.assert_called_once()
        self.assertEqual(base.read_bytes(), original)
        self.assertTrue((downloads / "261125.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
