# Consolidador de Notas de Crédito

Aplicación web para convertir PDFs de Notas de Crédito de Johnson Controls a Excel consolidado.

## 🚀 Características

- ✅ Sube múltiples PDFs a la vez
- ✅ Consolida productos duplicados automáticamente
- ✅ Calcula precios promedio y totales
- ✅ Genera archivos Excel con formato profesional
- ✅ Descarga todos los archivos en un ZIP

## 📦 Opciones de Hosting (GRATIS o Muy Económico)

### 1. **RENDER (RECOMENDADO) - GRATIS**

La opción más fácil y gratuita:

1. Ve a [render.com](https://render.com)
2. Crea una cuenta gratuita
3. Click en "New+" → "Web Service"
4. Conecta tu repositorio de GitHub
5. Configuración:
   - **Name**: `nc-consolidator`
   - **Environment**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Instance Type**: Free

¡Listo! Tu app estará en: `https://nc-consolidator.onrender.com`

**PROS**:
- ✅ Totalmente GRATIS
- ✅ SSL automático (HTTPS)
- ✅ Deploy automático desde GitHub
- ✅ No requiere tarjeta de crédito

**CONTRAS**:
- ⚠️ Se "duerme" después de 15 min sin uso (tarda ~30 seg en despertar)
- ⚠️ 750 horas/mes gratis

---

### 2. **RAILWAY - GRATIS (con límites)**

Otra excelente opción gratuita:

1. Ve a [railway.app](https://railway.app)
2. Crea cuenta con GitHub
3. Click "New Project" → "Deploy from GitHub repo"
4. Selecciona tu repositorio
5. Railway detecta Python automáticamente

URL: `https://tu-app.up.railway.app`

**PROS**:
- ✅ $5 USD de crédito GRATIS mensual
- ✅ No se duerme
- ✅ Muy rápido

**CONTRAS**:
- ⚠️ Después de $5/mes tienes que pagar
- ⚠️ Requiere tarjeta (pero no cobra si no pasas el límite)

---

### 3. **PYTHONANYWHERE - GRATIS**

Opción sencilla sin necesidad de Git:

1. Ve a [pythonanywhere.com](https://www.pythonanywhere.com)
2. Crea cuenta gratuita
3. Ve a "Web" → "Add a new web app"
4. Selecciona Flask
5. Sube tus archivos por FTP o desde consola

URL: `https://tuusuario.pythonanywhere.com`

**PROS**:
- ✅ Completamente GRATIS
- ✅ Fácil de usar
- ✅ No requiere conocimientos de Git

**CONTRAS**:
- ⚠️ Más lento
- ⚠️ Interfaz algo antigua

---

### 4. **VERCEL - GRATIS**

Requiere pequeñas modificaciones pero es muy rápido:

1. Ve a [vercel.com](https://vercel.com)
2. Importa desde GitHub
3. Vercel detecta Python automáticamente

**PROS**:
- ✅ Extremadamente rápido (CDN global)
- ✅ Totalmente GRATIS

**CONTRAS**:
- ⚠️ Serverless (límite de 10 seg por request)
- ⚠️ No ideal para procesamiento pesado

---

### 5. **GOOGLE CLOUD RUN - Casi GRATIS**

Para uso más intensivo:

```bash
gcloud run deploy nc-consolidator \
  --source . \
  --platform managed \
  --region us-central1 \
  --allow-unauthenticated
```

**PROS**:
- ✅ 2 millones de requests GRATIS/mes
- ✅ Escala automáticamente

**CONTRAS**:
- ⚠️ Requiere conocimientos técnicos
- ⚠️ Requiere tarjeta de crédito

---

## 🏆 MI RECOMENDACIÓN

**Para empezar: RENDER.COM**
- Es gratis
- No requiere tarjeta
- Funciona perfectamente para tu uso
- Deploy en 2 minutos

Si necesitas algo más rápido después: **RAILWAY**

---

## 💻 Instalación Local (Para Desarrollo)

```bash
# Instalar dependencias
pip install -r requirements.txt

# Ejecutar
python app.py

# Abrir en navegador
http://localhost:5000
```

---

## 📝 Cómo subir a GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/nc-consolidator.git
git push -u origin main
```

---

## 🔧 Variables de Entorno (Opcional)

Si quieres agregar configuraciones:

```
MAX_FILE_SIZE=50  # MB
DEBUG=False
```

---

## 📱 Uso de la Aplicación

1. Arrastra tus PDFs de Notas de Crédito
2. Click en "Procesar y Descargar Excel"
3. Descarga el ZIP con todos los archivos consolidados

---

## 🛠️ Tecnologías

- **Backend**: Flask (Python)
- **Procesamiento PDF**: pdfplumber
- **Generación Excel**: openpyxl
- **Frontend**: HTML/CSS/JavaScript

---

## 📄 Licencia

Uso libre para Energy Parts y Johnson Controls

---

## 👩‍💻 Soporte

Para dudas o mejoras, contacta a Claudia
