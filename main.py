"""Matriz de crecimiento-participación de la BCG a partir de matrizbcg.xlsx."""

import os
import sys
from pathlib import Path

import matplotlib

# PyCharm muestra los gráficos como imagen fija; en Mac abrimos una ventana
# real para que funcione el hover sobre los círculos.
if sys.platform == "darwin" and os.environ.get("MPLBACKEND", "").lower() not in ("agg", "pdf", "svg"):
    matplotlib.use("macosx")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from PIL import Image, ImageDraw, ImageFont

EXCEL = Path("/Users/karlasolarte/Andalucia FP/matrizbcg.xlsx")
SALIDA = Path(__file__).with_name("matriz_bcg.png")
SALIDA_WEB = Path(__file__).with_name("docs") / "index.html"  # publicada en GitHub Pages

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

# Color de cada círculo según su cuadrante
COLORES_CUADRANTE = {"Estrella": "#f2b705", "Interrogante": "#2a7de1",
                     "Vaca": "#2e9e5b", "Perro": "#c0392b"}


# Símbolo de cada cuadrante (emoji) y su alternativa si no hay fuente de emojis
FUENTE_EMOJI = "/System/Library/Fonts/Apple Color Emoji.ttc"
SIMBOLOS = {"Estrella": ("⭐", "★"), "Interrogante": ("❓", "?"),
            "Vaca": ("🐄", "$"), "Perro": ("🐕", "✕")}


def emoji_como_imagen(emoji):
    """Dibuja un emoji en color como imagen (matplotlib no sabe pintarlos como texto)."""
    fuente = ImageFont.truetype(FUENTE_EMOJI, 160)
    lienzo = Image.new("RGBA", (180, 180), (0, 0, 0, 0))
    ImageDraw.Draw(lienzo).text((0, 0), emoji, font=fuente, embedded_color=True)
    return np.asarray(lienzo.crop(lienzo.getbbox()))


def poner_simbolo(ax, cuadrante_, x, y, ha, va):
    """Coloca el símbolo del cuadrante en la esquina indicada."""
    emoji, alternativo = SIMBOLOS[cuadrante_]
    try:
        imagen = OffsetImage(emoji_como_imagen(emoji), zoom=0.22)
        caja = (0 if ha == "left" else 1, 1 if va == "top" else 0)
        ax.add_artist(AnnotationBbox(imagen, (x, y), box_alignment=caja, frameon=False))
    except OSError:  # sin fuente de emojis: símbolo de texto
        ax.text(x, y, alternativo, ha=ha, va=va, fontsize=20, color=GRIS, weight="bold")


def coma(valor, decimales=1):
    """Formato numérico español: 12.6 -> '12,6'."""
    return f"{valor:.{decimales}f}".replace(".", ",")


def dibujar(df):
    plt.rcParams["font.family"] = ["Arial", "DejaVu Sans"]

    fig, ax = plt.subplots(figsize=(11, 7.5))
    fig.subplots_adjust(left=0.08, right=0.95, top=0.95, bottom=0.10)

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

    # Símbolo de cada cuadrante en su esquina
    poner_simbolo(ax, "Estrella", x_max * 0.98, y_max, "left", "top")
    poner_simbolo(ax, "Interrogante", 0.02, y_max, "right", "top")
    poner_simbolo(ax, "Vaca", x_max * 0.98, y_min, "left", "bottom")
    poner_simbolo(ax, "Perro", 0.02, y_min, "right", "bottom")

    # 2. Círculos con el color de su cuadrante
    circulos = []
    for _, f in df.iterrows():
        area = f["peso_cartera_%"] * 80 + 150
        radio = area ** 0.5 / 2  # en puntos
        circulo = ax.scatter(f["cuota_relativa"], f["crecimiento_%"], s=area,
                             color=COLORES_CUADRANTE[f["cuadrante"]], alpha=0.85,
                             edgecolors="white", linewidths=1.5, zorder=2)
        circulos.append((circulo, f))

        # 3. Etiquetado directo (sin leyenda)
        ax.annotate(f["producto"], (f["cuota_relativa"], f["crecimiento_%"]),
                    textcoords="offset points", xytext=(radio + 6, 0),
                    ha="left", va="center", fontsize=9.5, color=GRIS_TEXTO, weight="bold")

    ax.set_xlim(x_max, 0)  # convención BCG: mayor cuota a la izquierda
    ax.set_ylim(y_min - 3, y_max + 3)

    # Títulos de eje centrados
    ax.set_xlabel("Cuota de mercado relativa", loc="center", color=GRIS_TEXTO, fontsize=10)
    ax.set_ylabel("Crecimiento del mercado", loc="center", color=GRIS_TEXTO, fontsize=10)

    fig.savefig(SALIDA, dpi=150)
    print(f"\nGráfico guardado en: {SALIDA}")

    activar_hover(fig, ax, circulos)
    plt.show()


def texto_hover(f):
    return (f"Cuota de mercado relativa: {coma(f['cuota_relativa'], 2)}\n"
            f"Crecimiento del mercado: {coma(f['crecimiento_%'])} %")


def activar_hover(fig, ax, circulos):
    """Muestra los datos de un producto al pasar el ratón por su círculo."""
    caja = ax.annotate("", xy=(0, 0), xytext=(15, 15), textcoords="offset points", fontsize=9.5,
                       color=GRIS_TEXTO, linespacing=1.5, zorder=10, multialignment="left",
                       bbox=dict(boxstyle="round,pad=0.6", facecolor="white", edgecolor=GRIS))
    caja.set_visible(False)

    def al_mover(evento):
        if evento.inaxes is not ax:
            tocados = []
        else:
            tocados = [(c, f) for c, f in circulos if c.contains(evento)[0]]
        if not tocados:
            if caja.get_visible():
                caja.set_visible(False)
                fig.canvas.draw_idle()
            return

        # Si dos círculos se solapan, el más cercano al ratón
        def distancia(par):
            x, y = ax.transData.transform(par[0].get_offsets()[0])
            return (x - evento.x) ** 2 + (y - evento.y) ** 2
        circulo, f = min(tocados, key=distancia)

        # La caja se abre hacia el centro del gráfico para que no se salga
        a_la_derecha = evento.x < fig.bbox.width / 2
        arriba = evento.y < fig.bbox.height / 2
        caja.xy = (f["cuota_relativa"], f["crecimiento_%"])
        caja.set_position((15 if a_la_derecha else -15, 15 if arriba else -15))
        caja.set_ha("left" if a_la_derecha else "right")
        caja.set_va("bottom" if arriba else "top")
        caja.set_text(texto_hover(f))
        caja.get_bbox_patch().set_edgecolor(COLORES_CUADRANTE[f["cuadrante"]])
        caja.set_visible(True)
        fig.canvas.draw_idle()

    fig.canvas.mpl_connect("motion_notify_event", al_mover)


# Cuadro de datos al pasar el ratón: se coloca a la izquierda del círculo (los
# nombres van a la derecha); si no cabe, encima, y si tampoco, debajo.
TOOLTIP_JS = """
const grafico = document.getElementById('{plot_id}');
const caja = document.createElement('div');
caja.style.cssText = 'position:absolute;display:none;pointer-events:none;background:white;'
  + 'border:1.5px solid;border-radius:6px;padding:6px 10px;font:13px Arial,Helvetica,sans-serif;'
  + 'color:#7F7F7F;line-height:1.5;white-space:nowrap;box-shadow:0 1px 4px rgba(0,0,0,.12);z-index:10';
grafico.parentElement.style.position = 'relative';
grafico.parentElement.appendChild(caja);

grafico.on('plotly_hover', (e) => {
  const p = e.points[0];
  const [cuota, crecimiento, color] = p.customdata;
  const x = p.xaxis.l2p(p.x) + p.xaxis._offset + grafico.offsetLeft;
  const y = p.yaxis.l2p(p.y) + p.yaxis._offset + grafico.offsetTop;
  const radio = p.data.marker.size[p.pointIndex] / 2;
  caja.innerHTML = `Cuota de mercado relativa: ${cuota}<br>Crecimiento del mercado: ${crecimiento} %`;
  caja.style.borderColor = color;
  caja.style.display = 'block';
  const ancho = caja.offsetWidth, alto = caja.offsetHeight;
  const izquierda = x - radio - 10 - ancho;
  let left, top;
  if (izquierda >= p.xaxis._offset + grafico.offsetLeft) {
    left = izquierda;
    top = y - alto / 2;
  } else {
    left = Math.min(Math.max(x - ancho / 2, 0), grafico.parentElement.clientWidth - ancho);
    top = (y - radio - 10 - alto >= 0) ? y - radio - 10 - alto : y + radio + 10;
  }
  caja.style.left = left + 'px';
  caja.style.top = top + 'px';
});
grafico.on('plotly_unhover', () => { caja.style.display = 'none'; });
"""


def crear_web(df):
    """Versión web interactiva del mismo gráfico (hover en el navegador)."""
    x_max = max(df["cuota_relativa"].max() * 1.2, 2 * UMBRAL_CUOTA)
    y_min = min(df["crecimiento_%"].min() - 10, -10)
    y_max = df["crecimiento_%"].max() + 15

    fig = go.Figure(go.Scatter(
        x=df["cuota_relativa"].tolist(), y=df["crecimiento_%"].tolist(),
        mode="markers+text", text=df["producto"], textposition="middle right",
        textfont=dict(color=GRIS_TEXTO, size=13, weight="bold"),
        marker=dict(size=((df["peso_cartera_%"] * 80 + 150) ** 0.5 * 1.33).tolist(), sizemode="diameter",
                    color=[COLORES_CUADRANTE[c] for c in df["cuadrante"]], opacity=0.85,
                    line=dict(color="white", width=1.5)),
        # El hover lo pinta TOOLTIP_JS (encima del círculo) para no tapar los nombres
        hoverinfo="none",
        customdata=[[coma(c, 2), coma(g), COLORES_CUADRANTE[q]]
                    for c, g, q in zip(df["cuota_relativa"], df["crecimiento_%"], df["cuadrante"])],
    ))

    # Líneas divisorias de la matriz
    linea = dict(color=GRIS_CLARO, width=1)
    fig.add_vline(x=UMBRAL_CUOTA, line=linea, layer="below")
    fig.add_hline(y=UMBRAL_CRECIMIENTO, line=linea, layer="below")

    # Símbolo de cada cuadrante en su esquina (eje X invertido)
    for emoji, x, y, xa, ya in [("⭐", 0.01, 0.99, "left", "top"), ("❓", 0.99, 0.99, "right", "top"),
                                ("🐄", 0.01, 0.01, "left", "bottom"), ("🐕", 0.99, 0.01, "right", "bottom")]:
        fig.add_annotation(text=emoji, x=x, y=y, xref="paper", yref="paper",
                           xanchor=xa, yanchor=ya, showarrow=False, font=dict(size=36))

    eje = dict(showgrid=False, zeroline=False, showline=True, linecolor=GRIS, ticks="",
               tickfont=dict(color=GRIS_TEXTO, size=12), title_font=dict(color=GRIS_TEXTO, size=13))
    marcas_x = np.arange(0, x_max + 0.01, 0.5)
    marcas_y = np.arange(np.ceil(y_min / 20) * 20, y_max + 0.01, 20)
    fig.update_xaxes(range=[x_max, 0], title="Cuota de mercado relativa",
                     tickvals=marcas_x.tolist(), ticktext=[coma(v) for v in marcas_x], **eje)
    fig.update_yaxes(range=[y_min - 3, y_max + 3], title="Crecimiento del mercado",
                     tickvals=marcas_y.tolist(), ticktext=[f"{v:+.0f} %".replace("+0", "0") for v in marcas_y], **eje)
    fig.update_layout(plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
                      font=dict(family="Arial, Helvetica, sans-serif"),
                      margin=dict(l=70, r=40, t=30, b=60), hovermode="closest")

    SALIDA_WEB.parent.mkdir(exist_ok=True)
    fig.write_html(SALIDA_WEB, include_plotlyjs="cdn", full_html=True,
                   config={"displayModeBar": False, "responsive": True},
                   default_width="100%", default_height="95vh", post_script=TOOLTIP_JS)
    html = SALIDA_WEB.read_text(encoding="utf-8")
    html = html.replace("<head>", '<head><title>Matriz BCG</title>'
                        '<meta name="viewport" content="width=device-width, initial-scale=1">', 1)
    SALIDA_WEB.write_text(html, encoding="utf-8")
    print(f"Página web guardada en: {SALIDA_WEB}")


def main():
    df = cargar_datos()
    tabla = df[["producto", "cuota_relativa", "crecimiento_%", "peso_cartera_%", "cuadrante"]]
    print(tabla.to_string(index=False, float_format=lambda v: f"{v:,.2f}"))
    crear_web(df)
    dibujar(df)


if __name__ == "__main__":
    main()
