"""
BAT 498 Final Project Source Code
Project: How Has AI Affected Unemployment?
Purpose: Clean annual BLS employment data by sector, compare 2014-2024 trends,
and generate the figures used in the final report/presentation.

Inputs expected:
  - bls_sector_employment.csv with columns: year, sector, employment
  - national_unemployment.csv with columns: year, unemployment_count or unemployment_rate
"""

import pandas as pd
import matplotlib.pyplot as plt

SECTOR_ORDER = [
    "Software Publishers",
    "Construction",
    "Management of Companies",
    "Schools",
    "Medical Equipment",
    "Administration Support Services",
]

PERIODS = {
    "Pre-COVID / Pre-AI": (2014, 2019),
    "COVID Disruption": (2020, 2021),
    "AI Period": (2022, 2024),
}


def clean_employment(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    df["sector"] = df["sector"].str.strip()
    df["year"] = pd.to_numeric(df["year"], errors="coerce").astype("Int64")
    df["employment"] = (
        df["employment"].astype(str).str.replace(",", "", regex=False).pipe(pd.to_numeric, errors="coerce")
    )
    return df.dropna(subset=["year", "sector", "employment"]).query("2014 <= year <= 2024")


def percent_change(df: pd.DataFrame) -> pd.Series:
    wide = df.pivot_table(index="year", columns="sector", values="employment", aggfunc="sum")
    return ((wide.loc[2024] - wide.loc[2014]) / wide.loc[2014] * 100).reindex(SECTOR_ORDER)


def annualized_growth(df: pd.DataFrame, start: int, end: int) -> pd.Series:
    wide = df.pivot_table(index="year", columns="sector", values="employment", aggfunc="sum")
    return (((wide.loc[end] / wide.loc[start]) ** (1 / (end - start))) - 1) * 100


def plot_employment_lines(df: pd.DataFrame, output_path: str) -> None:
    wide = df.pivot_table(index="year", columns="sector", values="employment", aggfunc="sum")
    ax = wide[SECTOR_ORDER].plot(figsize=(10, 6), linewidth=2)
    ax.set_title("Employment Rates by Sector, 2014-2024")
    ax.set_xlabel("Year")
    ax.set_ylabel("Employment")
    ax.grid(axis="y", alpha=0.30)
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)


def plot_percent_change(df: pd.DataFrame, output_path: str) -> None:
    changes = percent_change(df).sort_values()
    ax = changes.plot(kind="barh", figsize=(10, 5), color="#4e79a7")
    ax.set_title("Percent Employment Change by Sector, 2014-2024")
    ax.set_xlabel("Percent change")
    ax.grid(axis="x", alpha=0.25)
    for i, value in enumerate(changes):
        ax.text(value + 1, i, f"{value:.1f}%", va="center")
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)


def plot_period_growth(df: pd.DataFrame, output_path: str) -> None:
    result = pd.DataFrame({name: annualized_growth(df, *years) for name, years in PERIODS.items()})
    ax = result.reindex(SECTOR_ORDER).plot(kind="barh", figsize=(10, 6))
    ax.set_title("Employment Annualized Growth by Sector and Period")
    ax.set_xlabel("Annualized growth (%)")
    ax.axvline(0, color="black", linewidth=0.8)
    ax.grid(axis="x", alpha=0.25)
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)


if __name__ == "__main__":
    employment = clean_employment("bls_sector_employment.csv")
    plot_employment_lines(employment, "employment_rates_by_sector.png")
    plot_percent_change(employment, "percent_employment_change.png")
    plot_period_growth(employment, "annualized_growth_by_period.png")
    print(percent_change(employment).round(1))
