import unittest

import numpy as np
import pandas as pd

from calibrador import calcular, normalizar_pesos, pesos_iguales_porcentaje


class CalibradorTests(unittest.TestCase):
    def setUp(self):
        self.competencias = ["A", "B"]
        self.df = pd.DataFrame({
            "Candidato": ["Uno", "Dos"],
            "CAP_archivo": [80.0, 90.0],
            "A__valor": [8.0, 6.0],
            "A__esperado": [8.0, 8.0],
            "B__valor": [6.0, np.nan],
            "B__esperado": [8.0, 8.0],
        })
        self.esperados = {"A": 8.0, "B": 8.0}

    def test_pesos_iguales_reproducen_media(self):
        resultado = calcular(
            self.df, self.competencias, self.esperados, {"A": 1.0, "B": 1.0}
        )
        self.assertAlmostEqual(resultado.loc[0, "CAP_global"], 90.0)
        self.assertAlmostEqual(resultado.loc[1, "CAP_global"], 80.0)

    def test_ponderacion_cambia_el_cap(self):
        resultado = calcular(
            self.df, self.competencias, self.esperados, {"A": 3.0, "B": 1.0}
        )
        self.assertAlmostEqual(resultado.loc[0, "CAP_global"], 95.0)

    def test_escalar_todos_los_pesos_no_cambia_el_resultado(self):
        uno = calcular(self.df, self.competencias, self.esperados, {"A": 1, "B": 2})
        dos = calcular(self.df, self.competencias, self.esperados, {"A": 10, "B": 20})
        pd.testing.assert_series_equal(uno["CAP_global"], dos["CAP_global"])

    def test_rechaza_pesos_en_cero(self):
        with self.assertRaises(ValueError):
            normalizar_pesos(self.competencias, {"A": 0, "B": 0})

    def test_pesos_iniciales_suman_cien(self):
        pesos = pesos_iguales_porcentaje(["A", "B", "C", "D", "E", "F", "G"])
        self.assertAlmostEqual(sum(pesos.values()), 100.0)
        self.assertEqual(pesos["A"], 14.3)
        self.assertEqual(pesos["G"], 14.2)


if __name__ == "__main__":
    unittest.main()
