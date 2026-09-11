# Finanzas — App personal de gastos e ingresos

PWA (Progressive Web App) dark mode para iOS y Android.  
Se instala desde el navegador, funciona offline, datos en localStorage.

## Cómo publicar en GitHub Pages

### 1. Crear el repo en GitHub
1. Entrá a [github.com](https://github.com) → **New repository**
2. Nombre: `finanzas-app` (o el que quieras)
3. Visibilidad: **Public** (necesario para GitHub Pages gratis)
4. No agregues README ni .gitignore — el repo tiene que estar vacío

### 2. Subir el código
Desde la terminal, en la carpeta del proyecto:

```bash
git init
git add .
git commit -m "primera versión"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/finanzas-app.git
git push -u origin main
```

### 3. Activar GitHub Pages
1. En el repo → **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` / `/ (root)`
4. Guardar

En 1-2 minutos la app va a estar en:  
`https://TU_USUARIO.github.io/finanzas-app/`

### 4. Instalar en el celular

**iPhone (iOS Safari):**
1. Abrí la URL en Safari
2. Tocá el botón compartir (cuadrado con flecha)
3. "Agregar a inicio"

**Android / Samsung (Chrome):**
1. Abrí la URL en Chrome
2. Menú (tres puntos) → "Agregar a pantalla de inicio"
3. O el banner que aparece automáticamente

---

## Estructura del proyecto

```
finanzas-app/
├── index.html      ← estructura HTML
├── style.css       ← dark mode, diseño mobile
├── app.js          ← toda la lógica y datos
├── manifest.json   ← configuración PWA
├── sw.js           ← service worker (offline)
└── icons/
    ├── icon-192.png
    └── icon-512.png
```

## Categorías (rubros)
Hogar · Comida · Transporte · Entretenimiento · Salud · Moto · Salidas · Otros

## Métodos de pago (cajas)
MercadoPago · Amex · Visa Galicia · Master Galicia · Efectivo · Transferencia · Débito · Otros
