"""Verifica los cálculos de matrizbcgdatos.xlsx y dibuja la matriz BCG.

- Comprueba que todas las ventas son números en unidades completas.
- Recalcula con pandas la tasa de crecimiento, la cuota relativa, la proporción
  de la cartera y la posición en la matriz, y las compara con las fórmulas del
  Excel (si el archivo se ha abierto y guardado en Excel/Numbers).
- Dibuja la matriz con círculos proporcionales a la proporción de la cartera.
Todos los números se muestran en formato europeo (1.275.000 y coma decimal).
"""

from pathlib import Path

import matplotlib.pyplot as plt
import openpyxl
import pandas as pd

EXCEL = Path("/Users/karlasolarte/Andalucia FP/matrizbcgdatos.xlsx")
SALIDA = Path(__file__).with_name("matriz_bcg_verificada.png")

UMBRAL_CRECIMIENTO = 0.10  # 10 %
UMBRAL_CUOTA = 1.0
TOTAL_ESPERADO = 1_275_000

COLORES = {"Estrella": "#f2b705", "Interrogante": "#2a7de1",
           "Vaca lechera": "#2e9e5b", "Perro": "#c0392b"}


# ---------- Formato europeo ----------
def miles(valor):
    """1275000 -> '1.275.000'."""
    return f"{valor:,.0f}".replace(",", ".")


def decimal(valor, decimales=2):
    """0.88073 -> '0,88'."""
    return f"{valor:,.{decimales}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def porcentaje(fraccion, decimales=1):
    """0.12608 -> '12,6 %'."""
    return f"{decimal(fraccion * 100, decimales)} %"


# ---------- Cálculos ----------
def posicion(crecimiento, cuota):
    if crecimiento >= UMBRAL_CRECIMIENTO:
        return "Estrella" if cuota >= UMBRAL_CUOTA else "Interrogante"
    return "Vaca lechera" if cuota >= UMBRAL_CUOTA else "Perro"


def cargar_y_calcular():
    df = pd.read_excel(EXCEL, usecols="A:E")
    df.columns = ["producto", "ventas_empresa", "ventas_competidor",
                  "sector_actual", "sector_anterior"]

    # 1. Las ventas deben ser números puros (sin textos con puntos de miles)
    for col in df.columns[1:]:
        assert pd.api.types.is_numeric_dtype(df[col]), f"La columna {col} contiene texto"

    # 2. Total de ventas de la empresa
    total = df["ventas_empresa"].sum()
    assert total == TOTAL_ESPERADO, f"Total de ventas {miles(total)} ≠ {miles(TOTAL_ESPERADO)}"

    df["crecimiento"] = (df["sector_actual"] - df["sector_anterior"]) / df["sector_anterior"]
    df["cuota"] = df["ventas_empresa"] / df["ventas_competidor"]
    df["proporcion"] = df["ventas_empresa"] / total
    df["posicion"] = [posicion(g, c) for g, c in zip(df["crecimiento"], df["cuota"])]
    return df, total


def comparar_con_excel(df):
    """Compara con los valores que calcularon las fórmulas del Excel (columnas F:I)."""
    ws = openpyxl.load_workbook(EXCEL, data_only=True).active
    excel = [[ws[f"{c}{fila}"].value for c in "FGHI"] for fila in range(2, 2 + len(df))]
    if any(v is None for fila in excel for v in fila):
        print("\nAviso: el Excel aún no tiene los resultados de sus fórmulas guardados.\n"
              "Ábrelo en Excel o Numbers, guárdalo y vuelve a ejecutar para compararlos.")
        return
    for (_, f), (crec, cuota, prop, pos) in zip(df.iterrows(), excel):
        assert abs(f["crecimiento"] - crec) < 1e-9, f"Crecimiento distinto en {f['producto']}"
        assert abs(f["cuota"] - cuota) < 1e-9, f"Cuota distinta en {f['producto']}"
        assert abs(f["proporcion"] - prop) < 1e-9, f"Proporción distinta en {f['producto']}"
        assert f["posicion"] == pos, f"Posición distinta en {f['producto']}"
    print("\n✔ Las fórmulas del Excel coinciden con los cálculos de Python.")


def imprimir_tabla(df, total):
    tabla = pd.DataFrame({
        "Producto": df["producto"],
        "Ventas empresa": df["ventas_empresa"].map(lambda v: f"{miles(v)} €"),
        "Ventas competidor": df["ventas_competidor"].map(lambda v: f"{miles(v)} €"),
        "Tasa de crecimiento": df["crecimiento"].map(porcentaje),
        "Cuota relativa": df["cuota"].map(decimal),
        "Proporción cartera": df["proporcion"].map(porcentaje),
        "Posición BCG": df["posicion"],
    })
    print(tabla.to_string(index=False))
    print(f"\nTotal de ventas de la empresa: {miles(total)} €")


# ---------- Gráfico ----------
def dibujar(df, total):
    plt.rcParams["font.family"] = ["Arial", "DejaVu Sans"]
    fig, ax = plt.subplots(figsize=(12, 8))

    x_max = max(df["cuota"].max() * 1.25, 2 * UMBRAL_CUOTA)
    y_min = min(df["crecimiento"].min() - 0.10, -0.10)
    y_max = df["crecimiento"].max() + 0.15

    ax.axvline(UMBRAL_CUOTA, color="#BFBFBF", lw=1, zorder=0)
    ax.axhline(UMBRAL_CRECIMIENTO, color="#BFBFBF", lw=1, zorder=0)

    # Nombres de los cuadrantes (eje X invertido: más cuota a la izquierda)
    estilo = dict(fontsize=12, weight="bold", alpha=0.6)
    ax.text(x_max * 0.98, y_max * 0.97, "ESTRELLA", ha="left", va="top", color=COLORES["Estrella"], **estilo)
    ax.text(0.02, y_max * 0.97, "INTERROGANTE", ha="right", va="top", color=COLORES["Interrogante"], **estilo)
    ax.text(x_max * 0.98, y_min * 0.97, "VACA LECHERA", ha="left", va="bottom", color=COLORES["Vaca lechera"], **estilo)
    ax.text(0.02, y_min * 0.97, "PERRO", ha="right", va="bottom", color=COLORES["Perro"], **estilo)

    # Círculos: el ÁREA es proporcional a la proporción en la cartera
    escala = 9000  # puntos² para el 100 % de la cartera
    for _, f in df.iterrows():
        area = max(f["proporcion"] * escala, 40)  # mínimo para que el más pequeño se vea
        radio = area ** 0.5 / 2
        ax.scatter(f["cuota"], f["crecimiento"], s=area, color=COLORES[f["posicion"]],
                   alpha=0.85, edgecolors="white", linewidths=1.5, zorder=2)
        ax.annotate(f"{f['producto']}\n"
                    f"{miles(f['ventas_empresa'])} € · {porcentaje(f['proporcion'])} de la cartera",
                    (f["cuota"], f["crecimiento"]), textcoords="offset points",
                    xytext=(radio + 6, 0), ha="left", va="center", fontsize=9.5,
                    color="#595959", linespacing=1.4, zorder=3,
                    bbox=dict(facecolor="white", edgecolor="none", alpha=0.85, pad=1.5))

    ax.set_xlim(x_max, 0)
    ax.set_ylim(y_min, y_max)
    ax.xaxis.set_major_formatter(lambda v, _: decimal(v, 1))
    ax.yaxis.set_major_formatter(lambda v, _: porcentaje(v, 0))
    ax.set_xlabel("Cuota de mercado relativa (frente al principal competidor)", color="#595959")
    ax.set_ylabel("Tasa de crecimiento del mercado", color="#595959")
    ax.set_title(f"Matriz BCG · Ventas totales de la empresa: {miles(total)} €",
                 loc="left", fontsize=14, weight="bold", color="#404040")
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    for lado in ("left", "bottom"):
        ax.spines[lado].set_color("#BFBFBF")
    ax.tick_params(colors="#7F7F7F", length=0)

    fig.tight_layout()
    fig.savefig(SALIDA, dpi=150)
    print(f"\nGráfico guardado en: {SALIDA}")
    plt.show()


def main():
    df, total = cargar_y_calcular()
    imprimir_tabla(df, total)
    comparar_con_excel(df)
    dibujar(df, total)


if __name__ == "__main__":
    main()
