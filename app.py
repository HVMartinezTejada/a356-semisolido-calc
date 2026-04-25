"""
Calculadora de Proceso Semisólido — Aleación A356
Oh*(T): Número de Ohnesorge Modificado para Fabricación Aditiva en Estado Semisólido

Autores: Investigación en curso
Versión: 1.0
"""

import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from io import BytesIO
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Calculadora Semisólido A356 — Oh*(T)",
    page_icon="🔩",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CONSTANTES DEL SISTEMA A356
# ─────────────────────────────────────────────
PARAMS = {
    # Diagrama de fases Al-Si
    "Tl": 615.0,      # °C  — temperatura liquidus A356 (7% Si)
    "Te": 577.0,      # °C  — temperatura eutéctico Al-Si
    "C0": 7.0,        # wt% — composición inicial (Si)
    "ml": 6.79,       # °C/wt% — |pendiente liquidus| entre (7%,615°C) y (12.6%,577°C)
    "k":  0.11,       # adim — coeficiente de partición Al-Si

    # Reología (Krieger-Dougherty + Ley de Potencia)
    "eta0": 1.3e-3,   # Pa·s — viscosidad del Al líquido puro (~600°C)
    "fc":   0.60,     # adim — fracción sólida crítica
    "n":    2.0,      # adim — exponente K-D (forma de partícula)
    "m":    0.4,      # adim — índice de ley de potencia (shear-thinning, m<1)
    "g0":   1.0,      # s⁻¹  — tasa de corte de referencia

    # Propiedades físicas del A356 semisólido
    "rho":   2550.0,  # kg/m³ — densidad semisólida (interpolada sól/líq)
    "sigma": 0.860,   # N/m  — tensión superficial Al-Si líquido (~600°C)
}

# ─────────────────────────────────────────────
# MODELOS MATEMÁTICOS
# ─────────────────────────────────────────────
def fs_scheil(T, p=PARAMS):
    """
    Fracción sólida — Ecuación de Scheil (no equilibrio).
    Supone: sin difusión en sólido, mezcla completa en líquido.

    fs(T) = 1 - [1 + (Tl - T) / (|ml| · C0)]^(-1/(1-k))
    """
    dT = p["Tl"] - T
    dT = np.maximum(dT, 0.0)
    base = 1.0 + dT / (p["ml"] * p["C0"])
    exp_ = -1.0 / (1.0 - p["k"])
    fs = 1.0 - np.power(base, exp_)
    return np.clip(fs, 0.0, p["fc"] - 1e-4)


def fs_lever(T, p=PARAMS):
    """
    Fracción sólida — Regla de la Palanca (equilibrio).
    Supone: difusión completa en sólido y líquido.

    fs(T) = (Tl - T) / [(1-k) · (Tl - T + |ml| · C0)]
    """
    dT = p["Tl"] - T
    dT = np.maximum(dT, 0.0)
    denom = (1.0 - p["k"]) * (dT + p["ml"] * p["C0"])
    fs = np.where(denom > 1e-10, dT / denom, 0.0)
    return np.clip(fs, 0.0, p["fc"] - 1e-4)


def eta_app(fs, gamma_dot, p=PARAMS):
    """
    Viscosidad aparente — modelo combinado Krieger-Dougherty + Ley de Potencia.

    η = η₀ · (1 - fs/fc)^(-n) · (γ̇/γ̇₀)^(m-1)

    Unidades: Pa·s
    Válida para: 0 ≤ fs < fc,  γ̇ > 0
    """
    gamma_dot = np.maximum(gamma_dot, 1e-6)
    kd_term   = np.power(1.0 - fs / p["fc"], -p["n"])
    pl_term   = np.power(gamma_dot / p["g0"], p["m"] - 1.0)
    return p["eta0"] * kd_term * pl_term


def oh_star(fs, gamma_dot, D_m, p=PARAMS):
    """
    Número de Ohnesorge modificado para procesamiento semisólido.

    Oh*(T, γ̇, D) = η(fs, γ̇) / √(ρ · σ · D)

    Donde η es la viscosidad APARENTE (no dinámica newtoniana).
    Unidades: adimensional.
    D_m: diámetro de boquilla en metros.
    """
    eta_val = eta_app(fs, gamma_dot, p)
    denom   = np.sqrt(p["rho"] * p["sigma"] * D_m)
    return eta_val / denom


def power_thixo(fs, gamma_dot, p=PARAMS):
    """
    Potencia de thixoformado por unidad de volumen [W/m³]
    para mantener viscosidad constante η_c = η(fs=0, γ̇=γ̇₀).

    P(T) = P₀ · (1 - fs/fc)^(-2n/(1-m))
    P₀   = η_c · γ̇₀² · (η₀/η_c)^(2/(1-m))
    """
    eta_c = p["eta0"]    # viscosidad objetivo = viscosidad del líquido puro
    P0    = (eta_c * p["g0"]**2 *
             (p["eta0"] / eta_c) ** (2.0 / (1.0 - p["m"])))
    exp_  = -2.0 * p["n"] / (1.0 - p["m"])
    return P0 * np.power(1.0 - fs / p["fc"], exp_)


# ─────────────────────────────────────────────
# INTERPRETACIÓN DE ZONAS Oh*
# ─────────────────────────────────────────────
def zona_oh(oh_val):
    if oh_val < 0.1:
        return "⚡ Chorro / gotas (Oh* < 0.1)", "#1565C0"
    elif oh_val <= 10.0:
        return "✅ Depósito controlado (0.1 ≤ Oh* ≤ 10)", "#2E7D32"
    else:
        return "⛔ Demasiado viscoso para extruir (Oh* > 10)", "#E65100"


# ─────────────────────────────────────────────
# SIDEBAR — CONSOLA DE VARIABLES INDEPENDIENTES
# ─────────────────────────────────────────────
st.sidebar.title("🔩 Variables del Proceso")
st.sidebar.markdown("**Aleación:** A356 (Al-7%Si-0.3%Mg)")
st.sidebar.markdown("---")

st.sidebar.subheader("Geometría de boquilla")
D_mm = st.sidebar.slider(
    "Diámetro de boquilla D [mm]",
    min_value=0.5, max_value=15.0, value=3.0, step=0.5,
    help="Diámetro interno de la boquilla de extrusión. "
         "Afecta directamente Oh* vía √(ρ·σ·D)."
)
D_m = D_mm / 1000.0   # conversión a metros

st.sidebar.subheader("Condiciones de flujo")

gamma_exp = st.sidebar.slider(
    "γ̇ boquilla — mover para ajustar (escala log₁₀)",
    min_value=-2.0, max_value=3.0, value=1.0, step=0.1,
    format="%.1f",
    help="Tasa de corte característica en la boquilla. "
         "El slider controla el exponente de base 10. "
         "Valores típicos en AM semisólido: 10–100 s⁻¹ (exponente 1.0–2.0)."
)
gamma_dot = 10.0 ** gamma_exp
st.sidebar.metric(
    label="→ γ̇ boquilla [s⁻¹]",
    value=f"{gamma_dot:.3g} s⁻¹",
    help="Valor real de la tasa de corte en boquilla: γ̇ = 10^(slider)"
)

gamma_rest_exp = st.sidebar.slider(
    "γ̇ reposo post-depósito — mover para ajustar (log₁₀)",
    min_value=-3.0, max_value=0.0, value=-1.5, step=0.1,
    format="%.1f",
    help="Tasa de corte del material ya depositado. "
         "Valor bajo (~0.001–0.1 s⁻¹) representa reposo casi estático."
)
gamma_rest = 10.0 ** gamma_rest_exp
st.sidebar.metric(
    label="→ γ̇ reposo [s⁻¹]",
    value=f"{gamma_rest:.4g} s⁻¹",
    help="Valor real de la tasa de corte en reposo: γ̇ = 10^(slider)"
)

st.sidebar.subheader("Rango de temperatura")
T_min = st.sidebar.slider("T mínima [°C]", 578, 610, 580, step=1)
T_max = st.sidebar.slider("T máxima [°C]", 590, 614, 613, step=1)

st.sidebar.subheader("Modelo de fracción sólida")
modelo = st.sidebar.radio(
    "Seleccionar modelo",
    ["Ambos (Scheil + Palanca)", "Solo Scheil", "Solo Regla de la Palanca"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.markdown(
    f"""
**Parámetros del sistema:**
- η₀ = {PARAMS['eta0']*1000:.1f} mPa·s (Al líquido)
- f꜀ = {PARAMS['fc']} · n = {PARAMS['n']} · m = {PARAMS['m']}
- ρ = {PARAMS['rho']} kg/m³ · σ = {PARAMS['sigma']} N/m
- γ̇ boquilla = **{gamma_dot:.3g} s⁻¹**
- γ̇ reposo   = **{gamma_rest:.3g} s⁻¹**
- D = **{D_mm} mm**
    """
)

# ─────────────────────────────────────────────
# CÁLCULO DE VECTORES
# ─────────────────────────────────────────────
T_arr  = np.linspace(T_min, T_max, 400)
fs_s   = fs_scheil(T_arr)
fs_l   = fs_lever(T_arr)
eta_s  = eta_app(fs_s, gamma_dot)
eta_l  = eta_app(fs_l, gamma_dot)
oh_s   = oh_star(fs_s, gamma_dot, D_m)
oh_l   = oh_star(fs_l, gamma_dot, D_m)

# Estado en reposo (post-depósito)
oh_s_rest = oh_star(fs_s, gamma_rest, D_m)
oh_l_rest = oh_star(fs_l, gamma_rest, D_m)

pt_s   = power_thixo(fs_s, gamma_dot)
pt_l   = power_thixo(fs_l, gamma_dot)

# ─────────────────────────────────────────────
# CABECERA PRINCIPAL
# ─────────────────────────────────────────────
st.title("Calculadora de Proceso Semisólido — Aleación A356")
st.markdown(
    """
    **Herramienta de pre-diseño** para fabricación aditiva (AM) en estado semisólido.
    Integra los modelos de fracción sólida (Scheil / Regla de la Palanca),
    viscosidad aparente (Krieger-Dougherty + Ley de Potencia) y el
    **número de Ohnesorge modificado Oh\\*(T)** como criterio de ventana de proceso.
    """
)

# Tarjetas de estado
col1, col2, col3, col4 = st.columns(4)
oh_punto = oh_star(fs_scheil(np.array([590.0]))[0], gamma_dot, D_m)
oh_punto_rest = oh_star(fs_scheil(np.array([590.0]))[0], gamma_rest, D_m)
zona_txt, zona_color = zona_oh(oh_punto)

with col1:
    st.metric("γ̇ boquilla", f"{gamma_dot:.3g} s⁻¹")
with col2:
    st.metric("D boquilla", f"{D_mm} mm")
with col3:
    st.metric("Oh* @ 590°C (boquilla)", f"{oh_punto:.4f}")
with col4:
    st.metric("Oh* @ 590°C (reposo)", f"{oh_punto_rest:.4f}",
              delta=f"×{oh_punto_rest/oh_punto:.1f} vs boquilla")

st.markdown(f"> **Zona de flujo @ 590°C (boquilla):** {zona_txt}")
st.markdown("---")

# ─────────────────────────────────────────────
# TABS DE GRÁFICAS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📐 fₛ(T)",
    "💧 η(T) Viscosidad",
    "⭕ Oh\\*(T) — Principal",
    "⚡ P(T) Potencia",
    "ℹ️ Nota teórica Oh\\*",
])

# ── Colores consistentes ──
C_SCH   = "#1565C0"
C_LEV   = "#C62828"
C_REST  = "#1565C0"
C_REST2 = "#C62828"
ALPHA   = 0.75

def make_fig():
    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.tick_params(labelsize=10)
    return fig, ax

def add_legend_and_grid(ax):
    ax.grid(True, alpha=0.25, linestyle="--")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    leg = ax.legend(fontsize=9, framealpha=0.9)
    return leg

# ─── Tab 1: fₛ(T) ───────────────────────────
with tab1:
    st.subheader("Fracción sólida fₛ(T) — A356")
    st.markdown(
        "Evolución de la fracción sólida con la temperatura. "
        "La zona semisólida procesable para AM corresponde a **fₛ = 0.30–0.55**."
    )
    fig, ax = make_fig()
    if modelo != "Solo Regla de la Palanca":
        ax.plot(T_arr, fs_s, color=C_SCH, lw=2, label="Scheil (no equilibrio)")
    if modelo != "Solo Scheil":
        ax.plot(T_arr, fs_l, color=C_LEV, lw=2, ls="--", label="Regla de la Palanca (equilibrio)")

    ax.axhspan(0.30, 0.55, alpha=0.07, color="green")
    ax.axhline(0.30, color="green", ls=":", lw=1, alpha=0.6)
    ax.axhline(0.55, color="green", ls=":", lw=1, alpha=0.6)
    ax.text(T_min + 0.5, 0.42, "ventana AM (0.30–0.55)", fontsize=8,
            color="green", alpha=0.8)
    ax.axhline(PARAMS["fc"], color="gray", ls=":", lw=1, alpha=0.5)
    ax.text(T_min + 0.5, PARAMS["fc"] + 0.01, f"f꜀ = {PARAMS['fc']}", fontsize=8,
            color="gray")
    ax.set_xlabel("Temperatura T [°C]", fontsize=11)
    ax.set_ylabel("Fracción sólida fₛ  [adimensional]", fontsize=11)
    ax.set_ylim(0, 0.65)
    add_legend_and_grid(ax)
    st.pyplot(fig)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Scheil (no equilibrio):** sin difusión en sólido → mayor fₛ a igual T → estimación conservadora.")
    with col_b:
        st.markdown("**Regla de la Palanca (equilibrio):** difusión completa → menor fₛ → condición ideal, raramente alcanzada en práctica.")

# ─── Tab 2: η(T) ───────────────────────────
with tab2:
    st.subheader("Viscosidad aparente η(T) — A356")
    st.markdown(
        r"""
        **Modelo:** $\eta = \eta_0 \cdot (1 - f_s/f_c)^{-n} \cdot (\dot{\gamma}/\dot{\gamma}_0)^{m-1}$

        La viscosidad **no es constante** — depende de T (a través de fₛ) y de γ̇.
        Por eso se denomina *viscosidad aparente*: misma dimensión que la dinámica [Pa·s],
        pero su valor cambia con las condiciones del proceso.
        """
    )
    fig, ax = make_fig()
    if modelo != "Solo Regla de la Palanca":
        ax.semilogy(T_arr, eta_s, color=C_SCH, lw=2,
                    label=f"Scheil — boquilla γ̇={gamma_dot:.3g} s⁻¹")
        ax.semilogy(T_arr, eta_app(fs_s, gamma_rest), color=C_SCH, lw=2,
                    ls=":", alpha=0.6,
                    label=f"Scheil — reposo γ̇={gamma_rest:.3g} s⁻¹")
    if modelo != "Solo Scheil":
        ax.semilogy(T_arr, eta_l, color=C_LEV, lw=2, ls="--",
                    label=f"Palanca — boquilla γ̇={gamma_dot:.3g} s⁻¹")
        ax.semilogy(T_arr, eta_app(fs_l, gamma_rest), color=C_LEV, lw=2,
                    ls="-.", alpha=0.6,
                    label=f"Palanca — reposo γ̇={gamma_rest:.3g} s⁻¹")

    ax.axhline(PARAMS["eta0"], color="gray", ls=":", lw=1, alpha=0.5)
    ax.text(T_min + 0.5, PARAMS["eta0"] * 1.15,
            f"η₀ = {PARAMS['eta0']*1000:.1f} mPa·s (Al líquido puro)", fontsize=8, color="gray")
    ax.set_xlabel("Temperatura T [°C]", fontsize=11)
    ax.set_ylabel("Viscosidad aparente η [Pa·s]", fontsize=11)
    add_legend_and_grid(ax)
    st.pyplot(fig)

    st.info(
        "**Efecto shear-thinning:** la línea sólida (boquilla, γ̇ alto) siempre está "
        "por debajo de la punteada (reposo, γ̇ bajo). Esta diferencia es el mecanismo "
        "que hace viable la AM semisólida: fluye en la boquilla, retiene forma al depositarse."
    )

# ─── Tab 3: Oh*(T) — Principal ─────────────
with tab3:
    st.subheader("Número de Ohnesorge Modificado Oh\\*(T)")
    st.markdown(
        r"""
        $$Oh^*(T,\,\dot{\gamma},\,D) =
          \frac{\eta\bigl(f_s(T),\,\dot{\gamma}\bigr)}{\sqrt{\rho \cdot \sigma \cdot D}}$$

        **Variables independientes:** T, γ̇, D  
        **Variables dependientes:** fₛ(T) → η → Oh\\*  
        Escala logarítmica. Las líneas de zona marcan las fronteras del régimen de depósito.
        """
    )

    fig, ax = make_fig()

    # Zonas de color
    y_lo, y_hi = 1e-6, 1e3
    ax.fill_between(T_arr, y_lo, 0.1,  alpha=0.06, color="#1565C0")
    ax.fill_between(T_arr, 0.1,  10.0, alpha=0.06, color="#2E7D32")
    ax.fill_between(T_arr, 10.0, y_hi, alpha=0.06, color="#E65100")
    ax.axhline(0.1,  color="#1565C0", ls="--", lw=1.2, alpha=0.7)
    ax.axhline(10.0, color="#E65100", ls="--", lw=1.2, alpha=0.7)

    # Anotaciones de zona
    ax.text(T_arr[-1], 0.04,  "Chorro / gotas\n(Oh* < 0.1)",
            fontsize=8, color="#1565C0", ha="right", va="top")
    ax.text(T_arr[-1], 1.0,   "Depósito controlado\n(0.1 ≤ Oh* ≤ 10)",
            fontsize=8, color="#2E7D32", ha="right", va="center")
    ax.text(T_arr[-1], 30.0,  "Demasiado viscoso\n(Oh* > 10)",
            fontsize=8, color="#E65100", ha="right", va="bottom")

    if modelo != "Solo Regla de la Palanca":
        ax.semilogy(T_arr, oh_s, color=C_SCH, lw=2.5,
                    label=f"Scheil · boquilla (γ̇={gamma_dot:.3g} s⁻¹)")
        ax.semilogy(T_arr, oh_s_rest, color=C_SCH, lw=2, ls=":",
                    label=f"Scheil · reposo (γ̇={gamma_rest:.3g} s⁻¹)")
    if modelo != "Solo Scheil":
        ax.semilogy(T_arr, oh_l, color=C_LEV, lw=2.5, ls="--",
                    label=f"Palanca · boquilla (γ̇={gamma_dot:.3g} s⁻¹)")
        ax.semilogy(T_arr, oh_l_rest, color=C_LEV, lw=2, ls="-.",
                    label=f"Palanca · reposo (γ̇={gamma_rest:.3g} s⁻¹)")

    ax.set_xlabel("Temperatura T [°C]", fontsize=11)
    ax.set_ylabel("Oh*  [adimensional]", fontsize=11)
    ax.set_ylim(1e-5, 1e3)
    add_legend_and_grid(ax)
    st.pyplot(fig)

    # Tabla resumen a temperaturas clave
    st.markdown("#### Valores numéricos en temperaturas representativas")
    T_keys = [612, 606, 600, 593, 587, 582]
    rows = []
    for T_k in T_keys:
        if T_k < T_min or T_k > T_max:
            continue
        fs_sk = float(fs_scheil(np.array([T_k]))[0])
        fs_lk = float(fs_lever(np.array([T_k]))[0])
        rows.append({
            "T [°C]": T_k,
            "fₛ Scheil": f"{fs_sk:.3f}",
            "fₛ Palanca": f"{fs_lk:.3f}",
            f"η Scheil @ γ̇={gamma_dot:.2g} [Pa·s]":
                f"{float(eta_app(np.array([fs_sk]), gamma_dot)[0]):.4e}",
            f"Oh* Scheil boquilla":
                f"{float(oh_star(np.array([fs_sk]), gamma_dot, D_m)[0]):.4e}",
            f"Oh* Scheil reposo":
                f"{float(oh_star(np.array([fs_sk]), gamma_rest, D_m)[0]):.4e}",
            "Zona (boquilla)":
                zona_oh(float(oh_star(np.array([fs_sk]), gamma_dot, D_m)[0]))[0],
        })
    if rows:
        df_tabla = pd.DataFrame(rows)
        st.dataframe(df_tabla, use_container_width=True)

# ─── Tab 4: P(T) ────────────────────────────
with tab4:
    st.subheader("Potencia de thixoformado P(T) [W/m³]")
    st.markdown(
        r"""
        Potencia mecánica de agitación por unidad de volumen necesaria para
        mantener viscosidad constante η꜀ a lo largo del proceso de enfriamiento.

        $$P(T) = P_0 \cdot \left(1 - \frac{f_s(T)}{f_c}\right)^{-2n/(1-m)}$$

        Exponente: $-2n/(1-m) = -2 \times 2.0 / (1-0.4) \approx -6.67$
        """
    )
    fig, ax = make_fig()
    if modelo != "Solo Regla de la Palanca":
        ax.semilogy(T_arr, pt_s / 1e3, color=C_SCH, lw=2.5, label="Scheil")
    if modelo != "Solo Scheil":
        ax.semilogy(T_arr, pt_l / 1e3, color=C_LEV, lw=2.5, ls="--", label="Regla de la Palanca")
    ax.set_xlabel("Temperatura T [°C]", fontsize=11)
    ax.set_ylabel("P(T)  [kW/m³]", fontsize=11)
    add_legend_and_grid(ax)
    st.pyplot(fig)
    st.warning(
        "P(T) crece exponencialmente al acercarse fₛ → f꜀. "
        "La ventana práctica de operación se limita a fₛ < 0.55 para "
        "mantener requerimientos de potencia razonables."
    )

# ─── Tab 5: Nota teórica ────────────────────
with tab5:
    st.subheader("ℹ️ Significado y relevancia de Oh\\*(T)")
    st.markdown(
        r"""
        ### ¿Qué es el número de Ohnesorge clásico?

        El número de Ohnesorge **Oh** fue formulado originalmente para caracterizar
        la competencia entre fuerzas viscosas y las combinadas de inercia + tensión
        superficial en jets y gotas:

        $$Oh = \frac{\mu}{\sqrt{\rho \cdot \sigma \cdot L}}$$

        donde μ es la **viscosidad dinámica newtoniana** (constante para un fluido puro),
        ρ la densidad, σ la tensión superficial y L una longitud característica.

        ---

        ### ¿Por qué μ no basta para los semisólidos?

        La pasta semisólida de A356 **no es un fluido newtoniano**. Su viscosidad:

        - **Aumenta** cuando la fracción sólida fₛ crece (a menor temperatura).
        - **Disminuye** cuando la tasa de corte γ̇ aumenta (*shear-thinning*, m = 0.4 < 1).

        Usar μ constante en Oh equivaldría a ignorar ambos efectos, produciendo
        un único número que no refleja el estado real del proceso.

        ---

        ### La propuesta: Oh\\*(T)

        Se reemplaza μ por la **viscosidad aparente** η(fₛ, γ̇):

        $$Oh^*(T, \dot{\gamma}, D) =
          \frac{\eta_0 \cdot \left(1 - \dfrac{f_s(T)}{f_c}\right)^{-n}
          \cdot \left(\dfrac{\dot{\gamma}}{\dot{\gamma}_0}\right)^{m-1}}
          {\sqrt{\rho \cdot \sigma \cdot D}}$$

        donde fₛ(T) se obtiene mediante la ecuación de Scheil o la regla de la palanca.

        Oh\\*(T) ya **no es constante**: es una función del estado de proceso.

        ---

        ### Zonas de proceso para AM semisólida

        | Rango Oh\\* | Régimen | Consecuencia para AM |
        |---|---|---|
        | Oh\\* < 0.1 | Inercia + tensión superficial dominan | Formación de gotas/chorro inestable |
        | 0.1 ≤ Oh\\* ≤ 10 | Balance viscoso-capilar | **Depósito controlado y continuo** |
        | Oh\\* > 10 | Viscosidad domina completamente | Material no fluye por la boquilla |

        ---

        ### El papel del shear-thinning en AM

        El mecanismo que hace viable la AM semisólida es precisamente la diferencia entre
        Oh\\* en boquilla (γ̇ alto → η baja → Oh\\* bajo → fluye) y Oh\\* en reposo
        (γ̇ ≈ 0 → η alta → Oh\\* alto → retiene forma). Esta herramienta cuantifica
        ese contraste para cada temperatura y geometría de boquilla.

        ---

        ### Relevancia para la literatura

        La integración de:
        - El modelo de Krieger-Dougherty para fₛ(T), y
        - La ley de potencia para shear-thinning,

        dentro de la **forma funcional del número de Ohnesorge**, produce un criterio
        adimensional dependiente de la temperatura que **no existe explícitamente en la
        literatura para aleaciones Al-Si semisólidas**. Esta es la contribución original
        del artículo en preparación.

        **Referencias base del modelo:**
        - Einstein, A. (1906). Viscosidad de suspensiones diluidas.
        - Krieger & Dougherty (1959). Mecanismo de flujo no-newtoniano en suspensiones.
        - Flemings, M.C. (1991). Behavior of metal alloys in the semisolid state.
        - Spencer et al. (1972). Comportamiento reológico de Sn-15%Pb en solidificación.
        - Scheil, E. (1942). Remarks on the layer crystal formation.
        """
    )

# ─────────────────────────────────────────────
# EXPORTACIÓN A EXCEL
# ─────────────────────────────────────────────
st.markdown("---")
st.subheader("📥 Exportar datos para el artículo")
st.markdown(
    "Descarga un archivo Excel con todos los datos calculados para "
    "las condiciones actuales de la consola. Cada hoja contiene una variable, "
    "con ambos modelos (Scheil y Palanca) y columnas etiquetadas con unidades."
)

def build_excel(T_arr, fs_s, fs_l, eta_s, eta_l, oh_s, oh_l,
                oh_s_rest, oh_l_rest, pt_s, pt_l,
                gamma_dot, gamma_rest, D_mm, D_m):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Estilos ──
    hdr_fill   = PatternFill("solid", fgColor="1565C0")
    hdr_font   = Font(bold=True, color="FFFFFF", size=10)
    sub_fill   = PatternFill("solid", fgColor="E3F2FD")
    sub_font   = Font(bold=True, color="1565C0", size=10)
    note_font  = Font(italic=True, color="555555", size=9)
    center     = Alignment(horizontal="center")
    thin       = Side(style="thin", color="CCCCCC")
    brd        = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_sheet(ws, title, col_defs, data_cols, note=""):
        """col_defs: list of (header_str,), data_cols: list of arrays/lists.
        Handles both numeric and string columns correctly."""
        ws.title = title
        ws["A1"] = note
        ws["A1"].font = note_font
        ws.merge_cells(f"A1:{get_column_letter(len(col_defs))}1")
        for ci, (hdr,) in enumerate(col_defs, start=1):
            cell = ws.cell(row=2, column=ci, value=hdr)
            cell.fill = hdr_fill; cell.font = hdr_font
            cell.alignment = center; cell.border = brd
        for ri, vals in enumerate(zip(*data_cols), start=3):
            for ci, v in enumerate(vals, start=1):
                cell = ws.cell(row=ri, column=ci)
                cell.border = brd
                cell.alignment = center
                if isinstance(v, str):
                    cell.value = v
                else:
                    fv = float(v)
                    cell.value = round(fv, 8)
                    cell.number_format = (
                        "0.00000E+00"
                        if (fv != 0 and abs(fv) < 0.01)
                        else "0.000000"
                    )
        for ci in range(1, len(col_defs)+1):
            ws.column_dimensions[get_column_letter(ci)].width = 24

    # ── Hoja 1: fₛ(T) ──
    ws1 = wb.create_sheet()
    write_sheet(
        ws1, "fs(T)",
        [("T [°C]",), ("fs Scheil [adim]",), ("fs Palanca [adim]",)],
        [T_arr, fs_s, fs_l],
        note=f"Fracción sólida fₛ(T) — A356 (7% Si) | Tl={PARAMS['Tl']}°C | k={PARAMS['k']} | |ml|={PARAMS['ml']} °C/wt%"
    )

    # ── Hoja 2: η(T) ──
    ws2 = wb.create_sheet()
    write_sheet(
        ws2, "eta(T)",
        [("T [°C]",),
         (f"η Scheil boquilla [Pa·s] γ̇={gamma_dot:.3g}s⁻¹",),
         (f"η Palanca boquilla [Pa·s] γ̇={gamma_dot:.3g}s⁻¹",),
         (f"η Scheil reposo [Pa·s] γ̇={gamma_rest:.3g}s⁻¹",),
         (f"η Palanca reposo [Pa·s] γ̇={gamma_rest:.3g}s⁻¹",)],
        [T_arr, eta_s, eta_l,
         eta_app(fs_s, gamma_rest), eta_app(fs_l, gamma_rest)],
        note=(f"Viscosidad aparente η [Pa·s] | η₀={PARAMS['eta0']*1000:.1f}mPa·s | "
              f"fc={PARAMS['fc']} | n={PARAMS['n']} | m={PARAMS['m']}")
    )

    # ── Hoja 3: Oh*(T) ──
    ws3 = wb.create_sheet()
    write_sheet(
        ws3, "Oh_star(T)",
        [("T [°C]",),
         (f"Oh* Scheil boquilla [adim] D={D_mm}mm γ̇={gamma_dot:.3g}s⁻¹",),
         (f"Oh* Palanca boquilla [adim] D={D_mm}mm γ̇={gamma_dot:.3g}s⁻¹",),
         (f"Oh* Scheil reposo [adim] D={D_mm}mm γ̇={gamma_rest:.3g}s⁻¹",),
         (f"Oh* Palanca reposo [adim] D={D_mm}mm γ̇={gamma_rest:.3g}s⁻¹",),
         ("Zona proceso (Scheil boquilla)",)],
        [T_arr, oh_s, oh_l, oh_s_rest, oh_l_rest,
         [zona_oh(v)[0].split("(")[0].strip() for v in oh_s]],
        note=(f"Oh*(T) = η(fs,γ̇)/√(ρ·σ·D) | ρ={PARAMS['rho']}kg/m³ | "
              f"σ={PARAMS['sigma']}N/m | D={D_mm}mm | "
              f"Zona controlada: 0.1 ≤ Oh* ≤ 10")
    )

    # ── Hoja 4: P(T) ──
    ws4 = wb.create_sheet()
    write_sheet(
        ws4, "P(T)",
        [("T [°C]",), ("P Scheil [W/m³]",), ("P Palanca [W/m³]",)],
        [T_arr, pt_s, pt_l],
        note=(f"Potencia thixoformado P(T) [W/m³] | "
              f"Exponente = -2n/(1-m) = {-2*PARAMS['n']/(1-PARAMS['m']):.2f}")
    )

    # ── Hoja 5: Parámetros del sistema ──
    wsp = wb.create_sheet("Parametros")
    wsp["A1"] = "Parámetro"; wsp["A1"].fill = hdr_fill; wsp["A1"].font = hdr_font
    wsp["B1"] = "Valor";     wsp["B1"].fill = hdr_fill; wsp["B1"].font = hdr_font
    wsp["C1"] = "Unidades";  wsp["C1"].fill = hdr_fill; wsp["C1"].font = hdr_font
    wsp["D1"] = "Descripción"; wsp["D1"].fill = hdr_fill; wsp["D1"].font = hdr_font

    param_rows = [
        ("Tl",     PARAMS["Tl"],    "°C",      "Temperatura liquidus A356"),
        ("Te",     PARAMS["Te"],    "°C",      "Temperatura eutéctico Al-Si"),
        ("C0",     PARAMS["C0"],    "wt% Si",  "Composición inicial"),
        ("|ml|",   PARAMS["ml"],    "°C/wt%",  "Pendiente liquidus (valor absoluto)"),
        ("k",      PARAMS["k"],     "adim",    "Coeficiente de partición Al-Si"),
        ("η₀",     PARAMS["eta0"],  "Pa·s",    "Viscosidad Al líquido puro"),
        ("fc",     PARAMS["fc"],    "adim",    "Fracción sólida crítica"),
        ("n",      PARAMS["n"],     "adim",    "Exponente Krieger-Dougherty"),
        ("m",      PARAMS["m"],     "adim",    "Índice ley de potencia (shear-thinning)"),
        ("γ̇₀",    PARAMS["g0"],    "s⁻¹",     "Tasa de corte de referencia"),
        ("ρ",      PARAMS["rho"],   "kg/m³",   "Densidad semisólida A356"),
        ("σ",      PARAMS["sigma"], "N/m",     "Tensión superficial Al-Si líquido"),
        ("D",      D_mm,            "mm",      "Diámetro de boquilla (consola)"),
        ("γ̇ boquilla", gamma_dot,  "s⁻¹",     "Tasa de corte en boquilla (consola)"),
        ("γ̇ reposo",   gamma_rest, "s⁻¹",     "Tasa de corte post-depósito (consola)"),
    ]
    for ri, (p, v, u, d) in enumerate(param_rows, start=2):
        wsp.cell(ri, 1, p).border = brd
        wsp.cell(ri, 2, v).border = brd
        wsp.cell(ri, 3, u).border = brd
        wsp.cell(ri, 4, d).border = brd
    for ci in [1, 3, 4]:
        wsp.column_dimensions[get_column_letter(ci)].width = 22
    wsp.column_dimensions["B"].width = 15

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


col_dl1, col_dl2 = st.columns([1, 2])
with col_dl1:
    excel_buf = build_excel(
        T_arr, fs_s, fs_l, eta_s, eta_l, oh_s, oh_l,
        oh_s_rest, oh_l_rest, pt_s, pt_l,
        gamma_dot, gamma_rest, D_mm, D_m
    )
    st.download_button(
        label="⬇️ Descargar datos en Excel",
        data=excel_buf,
        file_name=f"A356_semisólido_D{D_mm}mm_gamma{gamma_dot:.1f}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col_dl2:
    st.markdown(
        f"""
        El archivo Excel contiene **5 hojas**:
        - **fs(T)** — fracción sólida Scheil y Palanca
        - **eta(T)** — viscosidad aparente en boquilla y reposo
        - **Oh_star(T)** — Oh\\* en boquilla y reposo, con columna de zona
        - **P(T)** — potencia de thixoformado
        - **Parametros** — tabla completa de constantes y variables de consola

        Condiciones exportadas: D = {D_mm} mm · γ̇ boquilla = {gamma_dot:.3g} s⁻¹ · γ̇ reposo = {gamma_rest:.3g} s⁻¹
        """
    )

st.markdown("---")
st.caption(
    "Calculadora de proceso semisólido A356 | "
    "Modelo reológico: Krieger-Dougherty + Ley de Potencia | "
    "Oh*(T): número de Ohnesorge modificado para AM semisólida"
)
