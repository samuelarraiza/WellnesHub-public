"""Generate a fully synthetic wellness export.

This repository ships no real data. The dashboard was originally built against a
club's internal wellness questionnaire, which contains named players and their
daily health answers. None of that is published here.

This script produces an Excel file with the exact same shape and column names as
that export, filled with invented players and randomly generated responses, so
the dashboard can be run and reviewed end to end.

    python scripts/generate_sample_data.py

Writes ``data/wellness_sample.xlsx`` by default.
"""

from __future__ import annotations

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

# Invented squad. Any resemblance to a real player is unintended.
PLAYERS = [
    "Alex Rivera", "Bruno Castell", "Carlos Nieto", "Diego Salas",
    "Erik Lindqvist", "Fabio Renzi", "Gabriel Costa", "Hugo Marchand",
    "Ivan Petrov", "Jonas Weber", "Kevin Byrne", "Matteo Ferrer",
    "Marco Bellini", "Nuno Cardoso", "Oscar Lindholm", "Ignasi Herrero",
    "Quim Batlle", "Ruben Solano", "Sergio Duarte", "Tomas Vidal",
    "Unai Beltran", "Victor Amaya",
]

COLUMNS = [
    "Fecha",
    "Jugador",
    "¿Qué tal has dormido?",
    "¿Cómo de fatigado te sientes?",
    "¿Tienes dolor muscular o de otro tipo?",
    "¿Te sientes estresado?",
    "¿Te apetece ir a entrenar?",
    "¿Cómo te encuentras animicamente?",
    "¿Ha sucedido algo fuera de lo normal que te gustaría comentar?",
    "¿Sientes que hay factores externos al club que te estén afectando últimamente?",
    "¿Cómo de fatigado te sientes? (POST)",
    "¿Tienes dolor muscular o de otro tipo? (POST)",
    "¿Te sientes estresado? (POST)",
    "¿Cómo te encuentras anímicamente? (POST)",
]

# Columns answered on a 1-5 Likert scale, and the direction that means "good".
# Higher is better for sleep, training desire and mood; lower is better for
# fatigue, muscle pain and stress.
POSITIVE_SCALE = {
    "¿Qué tal has dormido?": (3.6, 0.9),
    "¿Te apetece ir a entrenar?": (4.0, 0.8),
    "¿Cómo te encuentras animicamente?": (3.9, 0.8),
    "¿Cómo te encuentras anímicamente? (POST)": (3.8, 0.8),
}
NEGATIVE_SCALE = {
    "¿Cómo de fatigado te sientes?": (2.2, 0.9),
    "¿Tienes dolor muscular o de otro tipo?": (1.7, 0.9),
    "¿Te sientes estresado?": (1.8, 0.9),
    "¿Cómo de fatigado te sientes? (POST)": (2.9, 1.0),
    "¿Tienes dolor muscular o de otro tipo? (POST)": (2.1, 1.0),
    "¿Te sientes estresado? (POST)": (2.0, 0.9),
}


# Invented, deliberately unremarkable notes. The dashboard surfaces whatever is
# written here as an alert, so a handful of rows need text for that view to have
# anything to show.
FREE_TEXT_NOTES = [
    "Viaje largo el fin de semana",
    "He dormido mal, ruido en casa",
    "Molestia leve en el gemelo, nada que impida entrenar",
    "Semana de examenes",
    "Cambio de horario de entrenamiento",
]


def _likert(rng: np.random.Generator, mean: float, sd: float, bias: float) -> float:
    """Draw a 1-5 Likert answer around ``mean``, nudged by a per-player bias."""
    return float(np.clip(round(rng.normal(mean + bias, sd)), 1, 5))


def build(days: int, seed: int) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    random.seed(seed)

    # Each player has a stable temperament so the per-player views show variety
    # instead of everyone hovering around the same mean.
    bias = {p: rng.normal(0, 0.45) for p in PLAYERS}

    start = date.today() - timedelta(days=days)
    rows = []
    for offset in range(days):
        day = start + timedelta(days=offset)
        if day.weekday() == 6:  # rest day, no questionnaire
            continue
        # Not everyone answers every day.
        for player in [p for p in PLAYERS if rng.random() > 0.12]:
            row = {"Fecha": day.strftime("%d/%m/%Y"), "Jugador": player}
            b = bias[player]
            for col, (mean, sd) in POSITIVE_SCALE.items():
                row[col] = _likert(rng, mean, sd, b)
            for col, (mean, sd) in NEGATIVE_SCALE.items():
                row[col] = _likert(rng, mean, sd, -b)
            # The dashboard collapses these two free-text columns into a single
            # `alert` field, so a few rows carry a note to exercise that path.
            note = FREE_TEXT_NOTES[rng.integers(len(FREE_TEXT_NOTES))] if rng.random() < 0.05 else None
            row["¿Ha sucedido algo fuera de lo normal que te gustaría comentar?"] = note
            row["¿Sientes que hay factores externos al club que te estén afectando últimamente?"] = None
            rows.append(row)

    return pd.DataFrame(rows, columns=COLUMNS)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--days", type=int, default=120, help="days of history to generate")
    parser.add_argument("--seed", type=int, default=7, help="RNG seed, for reproducibility")
    parser.add_argument(
        "--out",
        type=Path,
        default=Path("data") / "wellness_sample.xlsx",
        help="destination .xlsx",
    )
    args = parser.parse_args()

    df = build(args.days, args.seed)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(args.out, index=False)
    print(f"{len(df)} rows for {df['Jugador'].nunique()} players -> {args.out}")


if __name__ == "__main__":
    main()
