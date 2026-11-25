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


# Paleta "Storytelling with Data": todo en gris y un solo color para el mensaje
GRIS_CLARO = "#D9D9D9"
GRIS = "#BFBFBF"
GRIS_TEXTO = "#7F7F7F"
GRIS_OSCURO = "#404040"
COLOR_FOCO = "#1F5FA6"


def coma(valor, decimales=1):
    """Formato numérico español: 12.6 -> '12,6'."""
    return f"{valor:.{decimales}f}".replace(".", ",")


def producto_foco(df):
    """Producto sobre el que gira la historia: el interrogante con más peso en ventas."""
    interrogantes = df[df["cuadrante"] == "Interrogante"]
    candidatos = interrogantes if not interrogantes.empty else df
    return candidatos.loc[candidatos["peso_cartera_%"].idxmax()]


def dibujar(df):
    plt.rcParams["font.family"] = ["Arial", "DejaVu Sans"]
    foco = producto_foco(df)

    fig, ax = plt.subplots(figsize=(11, 7.5))
    fig.subplots_adjust(left=0.08, right=0.95, top=0.80, bottom=0.12)

    x_max = max(df["cuota_relativa"].max() * 1.2, 2 * UMBRAL_CUOTA)
    y_min = min(df["crecimiento_%"].min() - 10, -10)
    y_max = df["crecimiento_%"].max() + 15

    # 1. Eliminar el desorden: sin rejilla, sin marco superior/derecho, ejes en gris
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    for lado in ("left", "bottom"):
        ax.spines[lado].set_color(GRIS)
    ax.tick_params(colors=GRIS_TEXTO, labelsize=10, length=0, pad=6)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:+.0f} %".replace("+0", "0"))
    ax.xaxis.set_major_formatter(lambda v, _: coma(v))

    # Divisiones de la matriz, discretas
    ax.axvline(UMBRAL_CUOTA, color=GRIS_CLARO, lw=1, zorder=0)
    ax.axhline(UMBRAL_CRECIMIENTO, color=GRIS_CLARO, lw=1, zorder=0)

    # Nombres de cuadrante en gris, sin competir con los datos
    cuad = dict(fontsize=11, color=GRIS, weight="bold")
    ax.text(x_max * 0.98, y_max, "ESTRELLAS", ha="left", va="top", **cuad)
    ax.text(0.02, y_max, "INTERROGANTES", ha="right", va="top", **cuad)
    ax.text(x_max * 0.98, y_min, "VACAS LECHERAS", ha="left", va="bottom", **cuad)
    ax.text(0.02, y_min, "PERROS", ha="right", va="bottom", **cuad)

    # 2. Todo en gris; solo el producto foco en color
    for _, f in df.iterrows():
        es_foco = f["producto"] == foco["producto"]
        area = f["peso_cartera_%"] * 80 + 150
        radio = area ** 0.5 / 2  # en puntos
        ax.scatter(f["cuota_relativa"], f["crecimiento_%"], s=area,
                   color=COLOR_FOCO if es_foco else GRIS,
                   edgecolors="white", linewidths=1.5, zorder=3 if es_foco else 2)

        # 3. Etiquetado directo (sin leyenda)
        if es_foco:
            texto = (f"{f['producto']}\n"
                     f"Cuota {coma(f['cuota_relativa'], 2)} · Mercado {f['crecimiento_%']:+.1f} %".replace(".", ","))
            ax.annotate(texto, (f["cuota_relativa"], f["crecimiento_%"]),
                        textcoords="offset points", xytext=(radio + 8, 0),
                        ha="left", va="center", fontsize=10.5, color=COLOR_FOCO,
                        weight="bold", linespacing=1.5,
                        bbox=dict(facecolor="white", edgecolor="none", pad=2))
        else:
            ax.annotate(f["producto"], (f["cuota_relativa"], f["crecimiento_%"]),
                        textcoords="offset points", xytext=(radio + 6, 0),
                        ha="left", va="center", fontsize=9.5, color=GRIS_TEXTO)

    ax.set_xlim(x_max, 0)  # convención BCG: mayor cuota a la izquierda
    ax.set_ylim(y_min - 3, y_max + 3)

    # Títulos de eje alineados arriba/izquierda para leerlos primero
    ax.set_xlabel("Cuota de mercado relativa", loc="left", color=GRIS_TEXTO, fontsize=10)
    ax.set_ylabel("Crecimiento del mercado", loc="top", color=GRIS_TEXTO, fontsize=10)

    # 4. Título que cuenta la conclusión, alineado a la izquierda
    if foco["cuadrante"] == "Interrogante":
        titulo = f"{foco['producto']} necesita una decisión de inversión"
        subtitulo = (f"Su mercado crece un {coma(foco['crecimiento_%'])} %, pero vendemos menos que el "
                     f"principal competidor. Supone el {coma(foco['peso_cartera_%'], 0)} % de nuestras ventas:\n"
                     "si ganamos cuota puede ser una estrella; si no, acabará como perro.")
    else:
        titulo = "Matriz BCG de la cartera de productos"
        subtitulo = ""
    fig.text(0.08, 0.95, titulo, fontsize=17, color=GRIS_OSCURO, weight="bold", ha="left", va="top")
    fig.text(0.08, 0.895, subtitulo, fontsize=11, color=GRIS_TEXTO, ha="left", va="top", linespacing=1.5)
    fig.text(0.08, 0.025, "Fuente: matrizbcg.xlsx", fontsize=8.5, color=GRIS, ha="left")

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
