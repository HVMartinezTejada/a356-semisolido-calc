# Calculadora de Proceso Semisólido — Aleación A356

**Herramienta de pre-diseño para fabricación aditiva (AM) en estado semisólido.**

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)

---

## Descripción

Esta aplicación implementa una calculadora interactiva para el **pre-diseño de procesos de fabricación aditiva con aleación A356 (Al-7%Si) en estado semisólido**. Integra tres modelos matemáticos interdependientes:

1. **Fracción sólida fₛ(T)** — Ecuación de Scheil (no equilibrio) y Regla de la Palanca (equilibrio)
2. **Viscosidad aparente η(T, γ̇)** — Modelo Krieger-Dougherty + Ley de Potencia (shear-thinning)
3. **Número de Ohnesorge modificado Oh\*(T)** — Criterio adimensional de ventana de proceso

La contribución central es la definición de **Oh\*(T)** como extensión del número de Ohnesorge clásico al dominio de fluidos no newtonianos en estado semisólido, integrando los parámetros de Krieger-Dougherty. Este número no ha sido derivado explícitamente en la literatura para aleaciones Al-Si semisólidas.

---

## Modelos matemáticos implementados

### Fracción sólida

**Scheil (no equilibrio):**
```
fs(T) = 1 - [1 + (Tl - T) / (|ml| · C₀)]^(-1/(1-k))
```

**Regla de la Palanca (equilibrio):**
```
fs(T) = (Tl - T) / [(1-k) · (Tl - T + |ml| · C₀)]
```

### Viscosidad aparente

```
η = η₀ · (1 - fs/fc)^(-n) · (γ̇/γ̇₀)^(m-1)
```

### Número de Ohnesorge modificado

```
Oh*(T, γ̇, D) = η(fs(T), γ̇) / √(ρ · σ · D)
```

### Potencia de thixoformado

```
P(T) = P₀ · (1 - fs(T)/fc)^(-2n/(1-m))
```

---

## Parámetros del sistema A356

| Parámetro | Valor | Unidad | Descripción |
|---|---|---|---|
| Tl | 615 | °C | Temperatura liquidus |
| Te | 577 | °C | Temperatura eutéctico Al-Si |
| C₀ | 7.0 | wt% Si | Composición inicial |
| \|ml\| | 6.79 | °C/wt% | Pendiente liquidus |
| k | 0.11 | adim | Coeficiente de partición |
| η₀ | 1.3×10⁻³ | Pa·s | Viscosidad Al líquido puro |
| fc | 0.60 | adim | Fracción sólida crítica |
| n | 2.0 | adim | Exponente Krieger-Dougherty |
| m | 0.4 | adim | Índice de ley de potencia |
| ρ | 2550 | kg/m³ | Densidad semisólida |
| σ | 0.86 | N/m | Tensión superficial Al-Si |

---

## Variables independientes (consola de la app)

| Variable | Rango | Unidad |
|---|---|---|
| D — diámetro de boquilla | 0.5 – 15 | mm |
| γ̇ en boquilla | 0.01 – 1000 | s⁻¹ |
| γ̇ en reposo (post-depósito) | 0.001 – 1.0 | s⁻¹ |
| Rango de temperatura | Tₑ+2 – Tₗ-1 | °C |
| Modelo fₛ | Scheil / Palanca / Ambos | — |

---

## Zonas de proceso (interpretación de Oh\*)

| Rango Oh\* | Régimen | Consecuencia para AM |
|---|---|---|
| Oh\* < 0.1 | Inercia + tensión superficial dominan | Formación de gotas / chorro inestable |
| 0.1 ≤ Oh\* ≤ 10 | Balance viscoso-capilar | **Depósito controlado y continuo** |
| Oh\* > 10 | Viscosidad domina | Material no fluye por la boquilla |

---

## Instalación local

```bash
git clone https://github.com/TU_USUARIO/a356-semisolido-calc.git
cd a356-semisolido-calc
pip install -r requirements.txt
streamlit run app.py
```

## Despliegue en Streamlit Cloud

1. Hacer fork o push de este repositorio a GitHub
2. Ir a [share.streamlit.io](https://share.streamlit.io)
3. Conectar repositorio → seleccionar `app.py` como archivo principal
4. Deploy

---

## Exportación de datos

La aplicación permite descargar un archivo Excel (`.xlsx`) con todas las variables calculadas para las condiciones actuales de la consola. El archivo contiene 5 hojas:

- **fs(T)** — fracción sólida Scheil y Palanca
- **eta(T)** — viscosidad aparente en boquilla y reposo
- **Oh_star(T)** — Oh\* en boquilla y reposo, con columna de zona de proceso
- **P(T)** — potencia de thixoformado
- **Parametros** — tabla completa de constantes del sistema

---

## Referencias

- Einstein, A. (1906). *Eine neue Bestimmung der Moleküldimensionen.*
- Krieger, I.M. & Dougherty, T.J. (1959). *A Mechanism for Non-Newtonian Flow in Suspensions of Rigid Spheres.*
- Flemings, M.C. (1991). *Behavior of Metal Alloys in the Semisolid State.* Metallurgical Transactions A.
- Spencer, D.B. et al. (1972). *Rheological Behavior of Sn-15%Pb in the Crystallization Range.*
- Scheil, E. (1942). *Bemerkungen zur Schichtkristallbildung.* Zeitschrift für Metallkunde.

---

## Artículo en preparación

Esta herramienta soporta el pre-diseño descrito en el artículo:

> *"Oh\*(T): Un número de Ohnesorge modificado para la caracterización de la ventana de proceso en fabricación aditiva de aleaciones Al-Si en estado semisólido"*
> (en preparación para revista Q1)
