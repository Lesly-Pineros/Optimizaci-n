import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

import asignacion_milp as base
import asignacion_tensores_milp as tensors
from salidas_asignacion import export_tobe, write_xlsx


class ModelRulesTest(unittest.TestCase):
    def setUp(self):
        self.project = {
            "Rol": "Dev Int", "Seniority": "Senior", "Dominio": "Operaciones",
            "Internal": "1", "External": "0",
            "_required_type": "interno",
            "Habilidades técnicas obligatorias 1": "Python",
            "Habilidades técnicas obligatorias 2": "Java",
        }
        self.person = {
            "key": "P1", "code": "P1", "name": "Persona", "role": "desarrollador",
            "role_raw": "Desarrollador", "seniority": "senior", "seniority_raw": "Senior",
            "company": "bbva", "resource_type": "interno", "direction": "ENG",
            "skills": {"python": 3, "java": 2},
        }
        self.role_map = {"dev int": "desarrollador"}

    def eligible(self, person=None, project=None):
        return base.eligibility(project or self.project, person or self.person, set(), self.role_map)

    def test_exact_seniority_is_required(self):
        person = dict(self.person, seniority="mid level")
        self.assertEqual(self.eligible(person=person)[2], "seniority")

    def test_domain_is_required_for_operations(self):
        person = dict(self.person, direction="")
        self.assertEqual(self.eligible(person=person)[2], "dominio")

    def test_internal_external_type_must_match(self):
        person = dict(self.person, resource_type="externo")
        self.assertEqual(self.eligible(person=person)[2], "tipo interno/externo")

    def test_dates_do_not_filter_initiatives(self):
        project = dict(self.project, **{"Fecha Inicio": "2Q 2027", "Fecha Fin": "4Q 2027"})
        self.assertEqual(len(tensors.make_slots([project])), 1)

    def test_lt_must_be_bbva_internal(self):
        project = dict(self.project, Rol="LT")
        person = dict(self.person, company="proveedor")
        role_map = {"lt": "desarrollador"}
        self.assertEqual(base.eligibility(project, person, set(), role_map)[2], "LT solo BBVA")

    def test_viability_matrix_prunes_invalid_pairs(self):
        slot = (2, dict(self.project, _required_type="interno"))
        people = [self.person, dict(self.person, key="P2", seniority="junior")]
        V, A, selected, stats = tensors.solve_lexicographic([slot], people, self.role_map)
        self.assertEqual(V.shape, (1, 2))
        self.assertEqual(int(V.sum()), 1)
        self.assertEqual(len(selected), 1)
        self.assertEqual(stats["coverage"], 1)
        self.assertEqual(stats["pruned_pct"], 50.0)

    def test_solver_uses_minimum_capacity_overflow_to_fill_all_slots(self):
        first = dict(self.project, **{"Nombre Proyecto": "P1", "SDATOOL": "1", "Mesa": "Mesa 1"})
        second = dict(self.project, **{"Nombre Proyecto": "P2", "SDATOOL": "2", "Mesa": "Mesa 2"})
        slots = [(2, dict(first)), (3, dict(second))]
        V, A, selected, stats = tensors.solve_lexicographic(slots, [self.person], self.role_map)
        self.assertEqual(len(selected), 2)
        self.assertEqual(stats["coverage"], 2)
        self.assertEqual(stats["capacity_overflow_people"], 1)
        baseline_slots, baseline_selected, _, _ = base.solve(
            [first, second], [self.person], set(), self.role_map)
        self.assertEqual(len(baseline_slots), 2)
        self.assertEqual(len(baseline_selected), 2)

    def test_xlsx_is_a_valid_zip_package(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "out.xlsx"
            write_xlsx(path, ["A", "B"], [{"A": "á", "B": "texto\n2"}])
            with ZipFile(path) as book:
                self.assertIn("xl/workbook.xml", book.namelist())
                self.assertIn("xl/worksheets/sheet1.xml", book.namelist())

    def test_tobe_exports_missing_mandatory_skills(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "tobe.xlsx"
            project = dict(self.project, **{"Fila origen": "2", "Nombre Proyecto": "P"})
            export_tobe(path, ["Fila origen", "Nombre Proyecto"], [project],
                        [(2, project)], [(0, 0, 0)], [self.person], lambda *_: (0, 2, 0, 2, 0),
                        base.TECH_FIELDS + base.FUNCTIONAL_FIELDS)
            with ZipFile(path) as book:
                sheet = book.read("xl/worksheets/sheet1.xml").decode()
                self.assertIn("Skills faltantes", sheet)
                self.assertIn("Java (nivel 2; requiere 3)", sheet)


if __name__ == "__main__":
    unittest.main()
