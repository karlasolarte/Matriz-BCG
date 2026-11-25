"""Matriz de crecimiento-participación de la BCG a partir de matrizbcg.xlsx."""

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

EXCEL = Path("/Users/karlasolarte/Andalucia FP/matrizbcg.xlsx")
SALIDA = Path(__file__).with_name("matriz_bcg.png")

# Umbrales de la matriz
UMBRAL_CUOTA = 1.0         # cuota relativa = 1 -> igual que el líder/competidor
UMBRAL_CRECIMIENTO = 10.0  # % de crecimiento del mercado


def a_numero(valor):
    """Convierte '2.590.000' (formato español) o 480.0 en float."""
    if isinstance(valor, str):
        valor = valor.strip().replace(".", "").replace(",", ".")
    return float(valor)


def cuadrante(cuota, crecimiento):
    if crecimiento >= UMBRAL_CRECIMIENTO:
        return "Estrella" if cuota >= UMBRAL_CUOTA else "Interrogante"
    return "Vaca" if cuota >= UMBRAL_CUOTA else "Perro"


def cargar_datos():
    df = pd.read_excel(EXCEL)
    df.columns = ["producto", "ventas_empresa", "ventas_competidor",
                  "sector_actual", "sector_anterior"]
    for col in df.columns[1:]:
        df[col] = df[col].map(a_numero)

    df["cuota_relativa"] = df["ventas_empresa"] / df["ventas_competidor"]
    df["crecimiento_%"] = (df["sector_actual"] - df["sector_anterior"]) / df["sector_anterior"] * 100
    df["peso_cartera_%"] = df["ventas_empresa"] / df["ventas_empresa"].sum() * 100
    df["cuadrante"] = [cuadrante(c, g) for c, g in zip(df["cuota_relativa"], df["crecimiento_%"])]
    return df


def dibujar(df):
    colores = {"Estrella": "#f2b705", "Interrogante": "#2a7de1",
               "Vaca": "#2e9e5b", "Perro": "#c0392b"}

    fig, ax = plt.subplots(figsize=(11, 8))

    x_max = max(df["cuota_relativa"].max() * 1.2, 2 * UMBRAL_CUOTA)
    y_min = min(df["crecimiento_%"].min() - 10, -10)
    y_max = df["crecimiento_%"].max() + 15

    # Líneas divisorias
    ax.axvline(UMBRAL_CUOTA, color="grey", lw=1.5)
    ax.axhline(UMBRAL_CRECIMIENTO, color="grey", lw=1.5)

    # Etiquetas de cuadrantes (convención BCG: cuota alta a la izquierda)
    etiqueta = dict(fontsize=15, alpha=0.3, weight="bold")
    ax.text(x_max * 0.98, y_max - 2, "ESTRELLA ★", ha="left", va="top", **etiqueta)
    ax.text(0.02, y_max - 2, "INTERROGANTE ?", ha="right", va="top", **etiqueta)
    ax.text(x_max * 0.98, y_min + 2, "VACA LECHERA $", ha="left", va="bottom", **etiqueta)
    ax.text(0.02, y_min + 2, "PERRO ✕", ha="right", va="bottom", **etiqueta)

    # Burbujas: tamaño proporcional a las ventas de la empresa
    for _, f in df.iterrows():
        area = f["peso_cartera_%"] * 80 + 150
        radio = area ** 0.5 / 2  # en puntos
        ax.scatter(f["cuota_relativa"], f["crecimiento_%"], s=area,
                   color=colores[f["cuadrante"]], alpha=0.7, edgecolors="black")
        ax.annotate(f"{f['producto']}\n(CR {f['cuota_relativa']:.2f} | TCM {f['crecimiento_%']:.1f}%)",
                    (f["cuota_relativa"], f["crecimiento_%"]),
                    textcoords="offset points", xytext=(radio + 6, 0),
                    ha="left", va="center", fontsize=8.5)

    ax.set_xlim(x_max, 0)  # eje invertido: mayor cuota a la izquierda
    ax.set_ylim(y_min, y_max)
    ax.set_xlabel("Cuota de mercado relativa")
    ax.set_ylabel("Tasa de crecimiento del mercado (%)")
    ax.set_title("Matriz BCG – Cartera de productos")
    ax.grid(alpha=0.2)

    fig.tight_layout()
    fig.savefig(SALIDA, dpi=150)
    print(f"\nGráfico guardado en: {SALIDA}")
    plt.show()


def main():
    df = cargar_datos()
    tabla = df[["producto", "cuota_relativa", "crecimiento_%", "peso_cartera_%", "cuadrante"]]
    print(tabla.to_string(index=False, float_format=lambda v: f"{v:,.2f}"))
    dibujar(df)


if __name__ == "__main__":
    main()
