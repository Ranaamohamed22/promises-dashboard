"""Protect the SNB Portfolio tab from running with only one input file.

Synthetic UI-control checks, no bank/customer data.
"""
import ast
import inspect
import textwrap
import unittest

from snb_performance import render_snb_performance_page


class TestPortfolioUploadGate(unittest.TestCase):
    def test_portfolio_tab_requires_both_uploaded_files(self):
        source = textwrap.dedent(inspect.getsource(render_snb_performance_page))
        tree = ast.parse(source)
        portfolio_with = next(
            node for node in ast.walk(tree)
            if isinstance(node, ast.With)
            and any(isinstance(context.context_expr, ast.Name)
                    and context.context_expr.id == "portfolio_tab"
                    for context in node.items)
        )
        guard = next(node for node in portfolio_with.body if isinstance(node, ast.If))
        compiled = compile(ast.Expression(body=guard.test), "<portfolio_upload_guard>", "eval")
        upload = object()
        for portfolio, payments, should_block in [
            (None, None, True),
            (upload, None, True),
            (None, upload, True),
            (upload, upload, False),
        ]:
            with self.subTest(portfolio=portfolio is not None,
                              payments=payments is not None):
                actual = eval(compiled, {"__builtins__": {}}, {
                    "portfolio_file": portfolio,
                    "payments_file": payments,
                })
                self.assertEqual(actual, should_block)

        # Portfolio processing must be in the ready/else branch, never the
        # branch reached by missing uploads.
        self.assertFalse(any(
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "render_portfolio_summary"
            for statement in guard.body for node in ast.walk(statement)
        ))
        self.assertTrue(any(
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "render_portfolio_summary"
            for statement in guard.orelse for node in ast.walk(statement)
        ))


if __name__ == "__main__":
    unittest.main()
