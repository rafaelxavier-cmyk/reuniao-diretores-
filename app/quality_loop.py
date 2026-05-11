"""
quality_loop.py – Loop de qualidade: gera, audita, ajusta, repete.

O loop roda até score >= 85 ou MAX_ITERATIONS vezes.
A cada iteração os parâmetros do gerador ficam mais conservadores.
"""
from __future__ import annotations
from pathlib import Path
from typing import Generator as GenType
import copy

MAX_ITERATIONS = 4

def run_quality_loop(
    pauta: dict,
    output_path: str,
    assets_dir: Path,
    progress_callback=None,
) -> tuple:
    """
    Gera, audita e melhora iterativamente o PPT.

    progress_callback(step: str, iteration: int, score: int | None) — opcional

    Retorna: (pptx_bytes, final_audit_report, iterations_log)
    """
    from generator import generate
    from layout_audit import audit

    iterations_log = []

    # Parâmetros que ficam progressivamente mais conservadores
    gen_params = {
        'max_bullets_per_card': 6,
        'card_title_lines':     2,
        'extra_spacing_factor': 1.0,
        'truncate_long_items':  False,
    }

    for iteration in range(1, MAX_ITERATIONS + 1):
        step = f"Iteracao {iteration}/{MAX_ITERATIONS}"
        if progress_callback:
            progress_callback(step, iteration, None)

        # Gera com os parâmetros atuais
        pauta_iter = _apply_params(pauta, gen_params)
        generate(pauta_iter, output_path, assets_dir=assets_dir)

        # Audita
        report = audit(output_path)

        iterations_log.append({
            'iteration': iteration,
            'score':     report.score,
            'issues':    len(report.issues),
            'params':    copy.deepcopy(gen_params),
        })

        if progress_callback:
            progress_callback(step, iteration, report.score)

        if report.passed:
            break

        # Ajusta parâmetros para a próxima iteração
        gen_params = _tighten_params(gen_params, report)

    # Auto-fix final com validator
    from validator import validate
    validate(output_path, autofix=True)

    pptx_bytes = Path(output_path).read_bytes()
    final_report = audit(output_path)

    return pptx_bytes, final_report, iterations_log


def _apply_params(pauta: dict, params: dict) -> dict:
    """
    Injeta os parâmetros de controle na pauta para que o gerador os use.
    O gerador lê estes de pauta['_gen_params'].
    """
    p = copy.deepcopy(pauta)
    p['_gen_params'] = params
    return p


def _tighten_params(params: dict, report) -> dict:
    """Torna os parâmetros mais conservadores com base nos issues."""
    p = copy.deepcopy(params)

    overflow_count = sum(1 for i in report.issues if i.kind == 'overflow')
    overlap_count  = sum(1 for i in report.issues if i.kind == 'overlap')
    tiny_count     = sum(1 for i in report.issues if i.kind == 'tiny_box')

    if overflow_count > 0 or tiny_count > 0:
        p['max_bullets_per_card'] = max(3, p['max_bullets_per_card'] - 1)
        p['extra_spacing_factor'] = min(1.5, p['extra_spacing_factor'] + 0.15)
        p['truncate_long_items']  = True

    if overlap_count > 0:
        p['card_title_lines']     = max(1, p['card_title_lines'] - 1)
        p['extra_spacing_factor'] = min(1.6, p['extra_spacing_factor'] + 0.1)

    return p
