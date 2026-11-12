import io
import os
import sys
import time
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.core.rounding_reconciler import (
    MAX_TOLERANCE_M,
    RoundingReconciler,
    largest_remainder_round,
)
from src.core.snapshot_vault_manager import SnapshotVaultManager
from src.core.validation_tie_out import ValidationTieOutEngine
from src.core.word_template_engine import WordTemplateEngine
from src.reporting.balance_generator import BalanceGenerator
from src.reporting.er_generator import ERGenerator


class TestRoundingReconciler(unittest.TestCase):
    """
    Suite automatizada T1–T5 para el Motor de Redondeo Jerárquico por Mayor Residuo
    y Reconciliación EEFF -> Notas (M$).
    """

    # =========================================================================
    # T1: Unidad — Algoritmo de Mayor Residuo (Largest Remainder / Hamilton)
    # =========================================================================
    def test_t1_largest_remainder_preserves_sum_and_zeros(self):
        raw_sub_thousand = {
            "Rubro A": 10_000_440.0,       # 10000.44
            "Rubro B": 20_000_460.0,       # 20000.46
            "Rubro C": 30_000_480.0,       # 30000.48
            "Rubro D": -5_000_530.0,       # -5000.53 -> floor -5001, rem 0.47
            "Rubro Cero": 0.0,             # Intocable en 0
            "Efectivo Fijo": 8_000_490.0,  # Bloqueado -> round(8000.49) = 8000
        }
        # Suma exacta = 63,001,340 pesos -> 63,001 M$, pero el redondeo individual da 62,999 M$ (-2 M$)
        naive_sum = sum(int(round(v / 1000.0)) for v in raw_sub_thousand.values())
        target_sub_m = int(round(sum(raw_sub_thousand.values()) / 1000.0))
        self.assertEqual(target_sub_m, 63001)
        self.assertEqual(naive_sum, 62999)

        reconciled, ok = largest_remainder_round(
            raw_sub_thousand,
            target_total_m=target_sub_m,
            scale_factor=1000.0,
            locked_keys={"Efectivo Fijo"},
        )
        self.assertTrue(ok)
        self.assertEqual(sum(reconciled.values()), target_sub_m)
        self.assertEqual(reconciled["Rubro Cero"], 0)
        self.assertEqual(reconciled["Efectivo Fijo"], 8000)
        self.assertEqual(reconciled["Rubro C"], 30001)
        self.assertEqual(reconciled["Rubro D"], -5000)

        for k, v in raw_sub_thousand.items():
            if v > 0:
                self.assertGreater(reconciled[k], 0)
            elif v < 0:
                self.assertLess(reconciled[k], 0)

    # =========================================================================
    # T2: Unidad — Guardarraíl Estricto de $4.000 Pesos (4 M$)
    # =========================================================================
    def test_t2_guardrail_blocks_discrepancy_over_4000_pesos(self):
        items = {
            "Cuenta 1": 50_000_000.0,
            "Cuenta 2": 30_000_000.0,
        }  # Suma = 80,000,000 (80.000 M$)

        # 1. Exactamente $4.000 pesos (4 M$) -> permitido
        res_ok, ok_4000 = largest_remainder_round(items, target_total_m=80004, scale_factor=1000.0)
        self.assertTrue(ok_4000)
        self.assertEqual(sum(res_ok.values()), 80004)
        self.assertEqual(res_ok["Cuenta 1"], 50004)

        # 2. 5 M$ ($5.000 > $4.000) -> BLOQUEADO
        res_blk, ok_5000 = largest_remainder_round(items, target_total_m=80005, scale_factor=1000.0)
        self.assertFalse(ok_5000)
        self.assertEqual(sum(res_blk.values()), 80000)

        # 3. $4.001 pesos en crudo -> BLOQUEADO
        items_4001 = {"Cuenta 1": 50_000_000.0, "Cuenta 2": 29_995_999.0}
        res_blk2, ok_4001 = largest_remainder_round(items_4001, target_total_m=80000, scale_factor=1000.0)
        self.assertFalse(ok_4001)
        self.assertEqual(sum(res_blk2.values()), 79996)

    # =========================================================================
    # T3: Integración EEFF — Cuadratura Vertical y Cruzada + Guardarraíl > $4.000
    # =========================================================================
    def test_t3_eeff_vertical_and_cross_statement_math(self):
        wte = WordTemplateEngine(None)

        # A) #ER con residuos fraccionarios acumulados
        pl_df = pd.DataFrame({
            "Ingresos de actividades ordinarias": [-100_000_440.0],
            "Costo de ventas": [40_000_560.0],
            "Gastos de administración": [15_000_550.0],
            "Costos financieros": [5_000_540.0],
            "Resultado por impuestos a las ganancias": [10_000_530.0],
            "Otros ingresos por función": [0.0],
        })
        expected_er_pesos = 100_000_440.0 - 40_000_560.0 - 15_000_550.0 - 5_000_540.0 - 10_000_530.0
        expected_er_m = int(round(expected_er_pesos / 1000.0))

        er_bytes, _ = ERGenerator("templates/Estado de Resultados Clasificados.xlsx").generate(
            pl_df, scale_factor=1000.0, periodo_actual_str="2099-10", periodo_comp_str="Ninguno"
        )
        df_er = wte._extract_er_preview(er_bytes, "2099-10", "Ninguno")
        col_er = df_er.columns[1]
        er_map = {
            str(r.iloc[0]).strip(): (0 if pd.isna(r[col_er]) else int(round(float(r[col_er]))))
            for _, r in df_er.iterrows()
            if pd.notna(r.iloc[0])
        }

        self.assertEqual(er_map["Otros ingresos por funcion"], 0)
        self.assertEqual(er_map["Ganancia Bruta"], er_map["Ingresos de actividades ordinarias"] + er_map["Costo de ventas"])
        self.assertEqual(
            er_map["Ganancia antes de impuesto"],
            er_map["Ganancia Bruta"] + er_map["Gastos de administración"] + er_map["Costos financieros"],
        )
        net_key = next(k for k in er_map if "del ejercicio" in k.lower())
        self.assertEqual(er_map[net_key], expected_er_m)

        # B) #BALANCE cuadrado en pesos
        tb_balanced = pd.DataFrame({
            "cuenta_id": ["101", "102", "103", "201", "202", "301", "302", "401"],
            "saldo_final": [
                15_000_490.0,     # Efectivo (bloqueado -> 15.000 M$)
                45_000_480.0,     # Deudores comerciales
                120_000_470.0,    # Propiedades, plantas y equipos
                -35_000_460.0,    # Cuentas comerciales por pagar, corrientes
                -55_000_450.0,    # Otros pasivos financieros, no corriente
                -50_000_000.0,    # Capital emitido
                -10_001_270.0,    # Resultados acumulados iniciales
                -29_999_260.0,    # Cuenta P&L
            ],
        })
        self.assertAlmostEqual(float(tb_balanced["saldo_final"].sum()), 0.0, places=4)

        map_bal = pd.DataFrame({
            "N° de Cuenta": ["101", "102", "103", "201", "202", "301", "302"],
            "Clasificación balance": [
                "Efectivo y efectivo equivalente",
                "Deudores comerciales y otras cuentas por cobrar, corrientes",
                "Propiedades, plantas y equipos",
                "Cuentas comerciales y otras cuentas por pagar, corrientes",
                "Otros pasivos financieros, no corriente",
                "Capital emitido",
                "Resultados acumulados",
            ],
        })

        bal_out = BalanceGenerator("templates/Balance clasificado.xlsx").generate(
            tb_balanced, map_bal, scale_factor=1000.0, periodo_actual_str="2099-10", periodo_comp_str="Ninguno"
        )
        df_bal = wte._extract_balance_preview(bal_out, "2099-10", "Ninguno")
        col_b = df_bal.columns[1]
        b_map = {
            str(r.iloc[0]).strip(): (0 if pd.isna(r[col_b]) else int(round(float(r[col_b]))))
            for _, r in df_bal.iterrows()
            if pd.notna(r.iloc[0])
        }

        self.assertEqual(b_map["Efectivo y efectivo equivalente"], 15000)
        self.assertEqual(
            b_map["Activos corrientes totales"],
            b_map["Efectivo y efectivo equivalente"] + b_map["Deudores comerciales y otras cuentas por cobrar, corrientes"],
        )
        self.assertEqual(b_map["Total activos"], b_map["Activos corrientes totales"] + b_map["Total activos no corrientes"])
        self.assertEqual(b_map["Total pasivos"], b_map["Pasivos corrientes totales"] + b_map["Total pasivos no corrientes"])
        self.assertEqual(b_map["Patrimonio total"], b_map["Capital emitido"] + b_map["Resultados acumulados"])
        self.assertEqual(b_map["Total activos"], b_map["Total patrimonio y pasivos"])

        # C) #BALANCE con descuadre real > $4.000 (+$15.000 pesos) -> Guardarraíl expone el descuadre
        tb_unbalanced = tb_balanced.copy()
        tb_unbalanced.loc[tb_unbalanced["cuenta_id"] == "102", "saldo_final"] += 15_000.0
        bal_bad_out = BalanceGenerator("templates/Balance clasificado.xlsx").generate(
            tb_unbalanced, map_bal, scale_factor=1000.0, periodo_actual_str="2099-10", periodo_comp_str="Ninguno"
        )
        df_bal_bad = wte._extract_balance_preview(bal_bad_out, "2099-10", "Ninguno")
        col_bb = df_bal_bad.columns[1]
        bb_map = {
            str(r.iloc[0]).strip(): (0 if pd.isna(r[col_bb]) else int(round(float(r[col_bb]))))
            for _, r in df_bal_bad.iterrows()
            if pd.notna(r.iloc[0])
        }
        self.assertGreaterEqual(abs(bb_map["Total activos"] - bb_map["Total patrimonio y pasivos"]), 14)

    # =========================================================================
    # T4: Notas — Líneas Protegidas, Matriz 2D, Cascada y Fórmulas Vivas =SUM(...)
    # =========================================================================
    def test_t4_note_reconciler_protected_rows_2d_and_cascade(self):
        wb = openpyxl.Workbook()

        ws_ef = wb.active
        ws_ef.title = "Efectivo"
        ws_ef.append(["#N04.1", None, None])
        ws_ef.append(["Concepto", "31.10.2026", "31.12.2025"])
        ws_ef.append(["Caja", 0, 0])                                  # R3: Cero
        ws_ef.append(["Efectivo en bancos", 100000, 90000])           # R4: Elegible -> absorbe +2 M$
        ws_ef.append(["Fondos mutuos", 5000000, 4000000])             # R5: Mayor absoluto pero PROTEGIDA por #N04.3
        ws_ef.append(["Total", "=SUM(B3:B5)", "=SUM(C3:C5)"])         # R6
        ws_ef.append([None, None, None])
        ws_ef.append(["#N04.2", None, None])
        ws_ef.append(["Moneda", "31.10.2026", "31.12.2025"])
        ws_ef.append(["Pesos chilenos", 4800000, 3800000])            # R10: Mayor en #N04.2 -> absorbe +3 M$
        ws_ef.append(["Dólares estadounidenses", 299999, 290000])     # R11
        ws_ef.append(["Total", "=SUM(B10:B11)", "=SUM(C10:C11)"])     # R12
        ws_ef.append([None, None, None])
        ws_ef.append(["#N04.3", None, None])
        ws_ef.append(["Fondo", "31.10.2026", "31.12.2025"])
        ws_ef.append(["Fondo Mutuo BCI", 5000000, 4000000])           # R16
        ws_ef.append(["Total", "=SUM(B16:B16)", "=SUM(C16:C16)"])     # R17

        ws_rou = wb.create_sheet("Activo por derechos de uso")
        ws_rou.append(["#N10.1", None, None, None])
        ws_rou.append([None, "31.10.2026", None, None])
        ws_rou.append(["Clase", "Valor bruto", "Depreciación acumulada", "Activo fijo neto"])
        ws_rou.append(["Inmuebles", 10000, -3000, "=SUM(B4:C4)"])      # R4
        ws_rou.append(["Vehículos", 25000, -8000, "=SUM(B5:C5)"])      # R5: Mayor depreciación -> absorbe -1 M$
        ws_rou.append(["Totales", "=SUM(B4:B5)", "=SUM(C4:C5)", "=SUM(D4:D5)"])  # R6

        ws_deu = wb.create_sheet("Deudores")
        ws_deu.append(["#N06.1", None, None])
        ws_deu.append(["Rubro", "31.10.2026", "31.12.2025"])
        ws_deu.append(["Clientes nacionales", 50000, 40000])
        ws_deu.append(["Total", "=SUM(B3:B3)", "=SUM(C3:C3)"])        # Diff = +12 M$ (> 4 M$ -> intocable)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        df_bal = pd.DataFrame({
            "Clasificación": [
                "Efectivo y efectivo equivalente",
                "Activo por derechos de uso",
                "Deudores comerciales y otras cuentas por cobrar, corrientes",
            ],
            "2026-10": [5100002, 23999, 50012],
        })

        raw_out, eval_out = RoundingReconciler.reconcile_and_evaluate_note_workbook(
            buf,
            empresa="Test SpA",
            periodo_actual="2026-10",
            scale_factor=1000.0,
            df_bal=df_bal,
            df_er=pd.DataFrame(),
        )

        wb_raw_res = openpyxl.load_workbook(raw_out, data_only=False)
        wb_eval_res = openpyxl.load_workbook(eval_out, data_only=False)

        self.assertEqual(wb_raw_res["Efectivo"]["B3"].value, 0)
        self.assertEqual(wb_raw_res["Efectivo"]["B5"].value, 5000000)
        self.assertEqual(wb_raw_res["Efectivo"]["B4"].value, 100002)
        self.assertEqual(wb_raw_res["Efectivo"]["B6"].value, "=SUM(B3:B5)")
        self.assertEqual(wb_eval_res["Efectivo"]["B6"].value, 5100002)
        self.assertEqual(wb_raw_res["Efectivo"]["B10"].value, 4800003)
        self.assertEqual(wb_raw_res["Efectivo"]["B12"].value, "=SUM(B10:B11)")
        self.assertEqual(wb_eval_res["Efectivo"]["B12"].value, 5100002)
        self.assertEqual(wb_eval_res["Efectivo"]["B17"].value, 5000000)

        self.assertEqual(wb_raw_res["Activo por derechos de uso"]["B5"].value, 25000)
        self.assertEqual(wb_raw_res["Activo por derechos de uso"]["C5"].value, -8001)
        self.assertEqual(wb_eval_res["Activo por derechos de uso"]["D5"].value, 16999)
        self.assertEqual(wb_eval_res["Activo por derechos de uso"]["D6"].value, 23999)

        self.assertEqual(wb_raw_res["Deudores"]["B3"].value, 50000)
        self.assertEqual(wb_eval_res["Deudores"]["B4"].value, 50000)

    # =========================================================================
    # T5: End-to-End Real — Períodos Reales, Inmutabilidad de Snapshots y Velocidad
    # =========================================================================
    def test_t5_end_to_end_real_periods_and_speed(self):
        # 1. Inmutabilidad de Snapshots cerrados en disco
        self.assertTrue(SnapshotVaultManager.has_active_snapshot("Pacifico Cable SpA", "2026-10"))
        _, hc_snap = ValidationTieOutEngine.obtener_matriz_tie_out("Pacifico Cable SpA", "2026-10", is_consolidated=False)
        self.assertEqual(hc_snap["source_mode"], "snapshot")

        # 2. Modo abierto / pre-cierre en vivo en las 3 entidades reales (0 diferencias de 1..4 M$)
        orig_has = SnapshotVaultManager.has_active_snapshot

        def _open_2026_10(emp, per):
            if str(per).strip() == "2026-10":
                return False
            return orig_has(emp, per)

        with patch.object(SnapshotVaultManager, "has_active_snapshot", side_effect=_open_2026_10):
            for emp, is_cons, min_real_big_diffs in [
                ("Pacifico Cable SpA", False, 10),
                ("Db Terra Chile Holdco SpA", False, 10),
                ("[GRUPO] Consolidado Terra Parent", True, 15),
            ]:
                RoundingReconciler.clear_cache()
                df_val, hc_val = ValidationTieOutEngine.obtener_matriz_tie_out(emp, "2026-10", is_consolidated=is_cons)
                self.assertEqual(hc_val["source_mode"], "live")

                small_rounding_diffs = df_val[
                    (df_val["Diferencia (M$)"] != 0) & (df_val["Diferencia (M$)"].abs() <= MAX_TOLERANCE_M)
                ]
                self.assertTrue(
                    small_rounding_diffs.empty,
                    f"Diferencias de redondeo (1..4 M$) remanentes en {emp}:\n{small_rounding_diffs}",
                )

                real_big_diffs = df_val[df_val["Diferencia (M$)"].abs() > MAX_TOLERANCE_M]
                self.assertGreaterEqual(len(real_big_diffs), min_real_big_diffs)
                self.assertTrue((real_big_diffs["Estado"] == "❌ Descuadre").all())

                pat_checks = df_val[df_val["Nota"].str.startswith("#PATRIMONIO")]
                self.assertFalse(pat_checks.empty)
                self.assertTrue(
                    (pat_checks["Diferencia (M$)"] == 0).all(),
                    f"Descuadre entre EEFF (#BALANCE / #ER / #PATRIMONIO) en {emp}:\n{pat_checks}",
                )

        # 3. Benchmark directo de sobrecosto del reconciliador (< 0.15 s)
        wb_bench = openpyxl.Workbook()
        ws_b = wb_bench.active
        ws_b.title = "Deudores"
        ws_b.append(["#N06.1", None])
        ws_b.append(["Concepto", "31.10.2026"])
        ws_b.append(["Clientes", 10000])
        ws_b.append(["Total", "=SUM(B3:B3)"])
        wb_eval_b = openpyxl.Workbook()
        ws_eb = wb_eval_b.active
        ws_eb.title = "Deudores"
        ws_eb.append(["#N06.1", None])
        ws_eb.append(["Concepto", "31.10.2026"])
        ws_eb.append(["Clientes", 10000])
        ws_eb.append(["Total", 10000])
        df_b = pd.DataFrame({
            "Clasificación": ["Deudores comerciales y otras cuentas por cobrar, corrientes"],
            "2026-10": [10001],
        })

        t_start = time.perf_counter()
        changed = RoundingReconciler.reconcile_note_workbooks_inplace(
            wb_raw=wb_bench, wb_eval=wb_eval_b, df_bal=df_b, df_er=pd.DataFrame(), periodo_actual="2026-10"
        )
        elapsed = time.perf_counter() - t_start
        self.assertEqual(changed, 1)
        self.assertLess(elapsed, 0.15)


if __name__ == "__main__":
    unittest.main()
