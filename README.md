# Finanzas — App personal de gastos e ingresos

PWA dark mode para iOS y Android. Datos en localStorage + backup automático en Google Drive.

---

## 1. Configurar Google OAuth (necesario para Drive)

Necesitás crear un Client ID gratis en Google Cloud. Son 5 minutos, solo se hace una vez.

### Paso a paso:

1. Entrá a [console.cloud.google.com](https://console.cloud.google.com)
2. **Crear proyecto** → nombre: `Finanzas App`
3. Menú izquierdo → **APIs y servicios** → **Biblioteca**
   - Buscá `Google Drive API` → Habilitar
4. **APIs y servicios** → **Credenciales** → **Crear credenciales** → **ID de cliente OAuth**
5. Tipo de aplicación: **Aplicación web**
6. Nombre: `Finanzas PWA`
7. **Orígenes de JavaScript autorizados** → Agregar:
   ```
   https://TU_USUARIO.github.io
   http://localhost:3000
   ```
8. **URIs de redireccionamiento** → Agregar:
   ```
   https://TU_USUARIO.github.io/finanzas-app/
   ```
9. Copiar el **Client ID** (tiene el formato `xxx.apps.googleusercontent.com`)

### Pegar el Client ID en la app:

Abrí `drive.js` y reemplazá la línea:
```js
const GOOGLE_CLIENT_ID = 'TU_CLIENT_ID_AQUI.apps.googleusercontent.com';
```
con tu Client ID real.

---

## 2. Publicar en GitHub Pages

### Crear el repo:
1. [github.com](https://github.com) → **New repository**
2. Nombre: `finanzas-app` · Público · Sin README

### Subir el código:
```bash
cd finanzas-app
git init
git add .
git commit -m "v1 con Drive sync"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/finanzas-app.git
git push -u origin main
```

### Activar GitHub Pages:
Settings → Pages → Branch: `main` / `/ (root)` → Save

URL final: `https://TU_USUARIO.github.io/finanzas-app/`

---

## 3. Instalar en el celular

**iPhone (iOS Safari):**
1. Abrí la URL en Safari
2. Botón compartir → "Agregar a inicio"

**Android / Samsung (Chrome):**
1. Abrí la URL en Chrome
2. Menú → "Agregar a pantalla de inicio"

---

## Cómo funciona la sincronización

| Evento | Qué pasa |
|---|---|
| Primer toque en "Conectar" | Login Google → descarga Drive → merge → sube |
| Agregar un movimiento | Guarda local → sube a Drive automáticamente |
| Eliminar un movimiento | Guarda local → sube a Drive automáticamente |
| Abrir la app (con token activo) | Baja Drive → merge con local |
| Sin internet | Todo funciona offline, sube cuando vuelve la conexión |

El archivo en Drive se llama `finanzas-data.json` y queda en "Mi unidad".

---

## Estructura del proyecto

```
finanzas-app/
├── index.html      ← estructura y shell HTML
├── style.css       ← dark mode, diseño mobile-first
├── app.js          ← lógica, datos y render
├── drive.js        ← sincronización con Google Drive
├── manifest.json   ← config PWA (instalación)
├── sw.js           ← service worker (offline)
└── icons/
    ├── icon-192.png
    └── icon-512.png
```
