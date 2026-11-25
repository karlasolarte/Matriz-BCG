"""Verifica los cálculos de matrizbcgdatos.xlsx y dibuja la matriz BCG.

- Comprueba que todas las ventas son números en unidades completas.
- Recalcula con pandas la tasa de crecimiento, la cuota relativa, la proporción
  de la cartera y la posición en la matriz, y las compara con las fórmulas del
  Excel (si el archivo se ha abierto y guardado en Excel/Numbers).
- Dibuja la matriz con círculos proporcionales a la proporción de la cartera.
  Cada círculo lleva el nombre del producto; sus ventas y su proporción en la
  cartera aparecen al pasar el ratón por encima (en la ventana del gráfico).
Todos los números se muestran en formato europeo (1.275.000 y coma decimal).
"""

import os
import sys
from pathlib import Path

import matplotlib

# PyCharm muestra los gráficos como imagen fija; en Mac abrimos una ventana
# real para que funcione el hover sobre los círculos.
if sys.platform == "darwin" and os.environ.get("MPLBACKEND", "").lower() not in ("agg", "pdf", "svg"):
    matplotlib.use("macosx")

import matplotlib.pyplot as plt
import openpyxl
import pandas as pd
from matplotlib.transforms import Bbox

EXCEL = Path(__file__).with_name("matrizbcgdatos.xlsx")  # el Excel va junto al script
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
def dibujar(df):
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
    circulos, nombres = [], []
    for _, f in df.iterrows():
        area = max(f["proporcion"] * escala, 40)  # mínimo para que el más pequeño se vea
        radio = area ** 0.5 / 2
        circulo = ax.scatter(f["cuota"], f["crecimiento"], s=area, color=COLORES[f["posicion"]],
                             alpha=0.85, edgecolors="white", linewidths=1.5, zorder=2)
        # Solo el nombre; los datos se ven al pasar el ratón
        nombre = ax.annotate(f["producto"], (f["cuota"], f["crecimiento"]), textcoords="offset points",
                             xytext=(radio + 6, 0), ha="left", va="center", fontsize=10,
                             weight="bold", color="#595959", zorder=3,
                             bbox=dict(facecolor="white", edgecolor="none", alpha=0.85, pad=1.5))
        circulos.append((circulo, f, radio))
        nombres.append(nombre)

    ax.set_xlim(x_max, 0)
    ax.set_ylim(y_min, y_max)
    ax.xaxis.set_major_formatter(lambda v, _: decimal(v, 1))
    ax.yaxis.set_major_formatter(lambda v, _: porcentaje(v, 0))
    ax.set_xlabel("Cuota de mercado relativa (frente al principal competidor)", color="#595959")
    ax.set_ylabel("Tasa de crecimiento del mercado", color="#595959")
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    for lado in ("left", "bottom"):
        ax.spines[lado].set_color("#BFBFBF")
    ax.tick_params(colors="#7F7F7F", length=0)

    fig.tight_layout()
    fig.savefig(SALIDA, dpi=150)
    print(f"\nGráfico guardado en: {SALIDA}")

    activar_hover(fig, ax, circulos, nombres)
    plt.show()


def activar_hover(fig, ax, circulos, nombres):
    """Al pasar el ratón por un círculo muestra sus ventas y su proporción en la
    cartera, en la posición libre más cercana que no tape ningún nombre."""
    tamano_letra = 10
    caja = ax.annotate("", xy=(0, 0), xycoords="figure pixels", ha="left", va="bottom",
                       fontsize=tamano_letra, color="#595959", linespacing=1.5, zorder=10,
                       annotation_clip=False,
                       bbox=dict(boxstyle="round,pad=0.6", facecolor="white", edgecolor="#BFBFBF"))
    caja.set_visible(False)

    def solape(a, b):
        return (max(0, min(a.x1, b.x1) - max(a.x0, b.x0)) *
                max(0, min(a.y1, b.y1) - max(a.y0, b.y0)))

    def al_mover(evento):
        tocados = []
        if evento.inaxes is ax:
            tocados = [(c, f, r) for c, f, r in circulos if c.contains(evento)[0]]
        if not tocados:
            if caja.get_visible():
                caja.set_visible(False)
                fig.canvas.draw_idle()
            return

        # Si dos círculos se solapan, el más cercano al ratón
        def distancia(t):
            x, y = ax.transData.transform(t[0].get_offsets()[0])
            return (x - evento.x) ** 2 + (y - evento.y) ** 2
        circulo, f, radio_pt = min(tocados, key=distancia)

        caja.set_text(f"Ventas: {miles(f['ventas_empresa'])} €\n"
                      f"Proporción de la cartera: {porcentaje(f['proporcion'])}")
        caja.get_bbox_patch().set_edgecolor(COLORES[f["posicion"]])
        caja.set_visible(True)  # oculta, matplotlib la mediría con tamaño cero

        # Medidas en píxeles de pantalla
        render = fig.canvas.get_renderer()
        px = fig.dpi / 72
        relleno = 0.6 * tamano_letra * px
        texto = caja.get_window_extent(render)
        w, h = texto.width + 2 * relleno, texto.height + 2 * relleno
        cx, cy = ax.transData.transform((f["cuota"], f["crecimiento"]))
        r, sep = radio_pt * px, 8 * px
        zona = ax.get_window_extent(render)
        textos = [n.get_window_extent(render) for n in nombres]
        mi_nombre = textos[[t[1]["producto"] for t in circulos].index(f["producto"])]

        # Posiciones candidatas (esquina inferior izquierda), en orden de preferencia
        candidatas = [
            (cx - r - sep - w, cy - h / 2),        # a la izquierda del círculo
            (cx - w / 2, cy + r + sep),            # encima
            (cx - w / 2, cy - r - sep - h),        # debajo
            (cx - r - sep - w, cy + r + sep),      # arriba a la izquierda
            (cx - r - sep - w, cy - r - sep - h),  # abajo a la izquierda
            (mi_nombre.x1 + sep, cy - h / 2),      # a la derecha del nombre
        ]
        mejor = None
        for orden, (x0, y0) in enumerate(candidatas):
            x0 = min(max(x0, zona.x0), zona.x1 - w)
            y0 = min(max(y0, zona.y0), zona.y1 - h)
            rect = Bbox.from_bounds(x0, y0, w, h)
            coste = orden + sum(solape(rect, t) for t in textos) * 100
            for c, _, rr in circulos:
                bx, by = ax.transData.transform(c.get_offsets()[0])
                rr *= px
                coste += solape(rect, Bbox.from_extents(bx - rr, by - rr, bx + rr, by + rr))
            if mejor is None or coste < mejor[0]:
                mejor = (coste, x0, y0)

        _, x0, y0 = mejor
        caja.xy = (x0 + relleno, y0 + relleno)
        caja.set_position(caja.xy)
        caja.set_visible(True)
        fig.canvas.draw_idle()

    fig.canvas.mpl_connect("motion_notify_event", al_mover)


def main():
    df, total = cargar_y_calcular()
    imprimir_tabla(df, total)
    comparar_con_excel(df)
    dibujar(df)


if __name__ == "__main__":
    main()
