# 🛡️ Peña Seguro 1410 — Plataforma de Inspectores Municipales

Aplicación web progresiva (Mobile-First) para la gestión operativa en terreno de **Peña Seguro 1410 (Ilustre Municipalidad de Peñaflor)**, conectada en tiempo real con **Google Sheets (Excel en Google Drive)** y **Google Drive (Carpeta de Evidencias Fotográficas)** mediante **Google Apps Script**.

---

## 📂 Estructura del Repositorio

```text
├── index.html                        # Aplicación principal para GitHub Pages (Inspectores)
├── pe_a_seguro_inspectores.html      # Plantilla fuente principal sincronizada
├── sistema_de_reportes_operativos_pe_a_seguro_1410(1).html # Generador de reportes WhatsApp/Correo
├── google-apps-script/
│   └── Code.gs                       # Backend gratuito en Google Apps Script (Sheets + Drive + Gmail)
├── .gitignore
└── README.md
```

---

## ⚙️ Arquitectura: ¿Cómo se registra todo en un Excel en Google Drive?

Dado que **GitHub Pages** aloja sitios estáticos (`HTML/JS`) de forma 100% gratuita y rápida, la forma más segura, gratuita y sin servidores intermedios para guardar datos y fotografías en Google Drive es utilizar un **Web App de Google Apps Script (`Code.gs`)** vinculado a una Hoja de Cálculo de Google:

1. **Celular del Inspector (`index.html`)**:
   - Captura **Apertura de Turno**, **Checklist Técnico Exhaustivo** (24 puntos para Móviles/Camionetas o Motocicletas), **Procedimientos**, **Visto Bueno de Tareas del Cuadrante** y **Cierre de Turno (Kilometraje)**.
   - Comprime automáticamente las fotografías en el navegador (`canvas` JPEG a máx. 1024px) para que suban en 1–2 segundos incluso con señal móvil débil.
   - Si no hay señal en terreno, guarda el registro en una **Cola Offline (`localStorage`)** y lo sube automáticamente cuando vuelve la conexión (también permite descargar un respaldo `.CSV` compatible con Excel en cualquier momento).
2. **Puente Serverless (`google-apps-script/Code.gs`)**:
   - Recibe las peticiones `POST` desde la aplicación web.
   - Guarda las fotografías en una carpeta de Google Drive (`Peña_Seguro_1410_Evidencias`) y genera el enlace directo de visualización.
   - Inserta las filas automáticamente en **14 pestañas** dentro del archivo de Google Sheets:
     - **Inspectores & Flota (existentes)**: `Turnos_y_Checklist`, `Procedimientos`, `Tareas_Cuadrante`, `Cierres_Turno`, `Usuarios`, `Catalogo_Tareas`, `Config_Flota`, `Mantenciones`, `Alertas_Flota`.
     - **Central CENCO CAD 1410 (nuevas)**:
       - `CENCO_Bitacora_CAD` (Bitácora unificada con imágenes en Drive + espejo automático de procedimientos de Inspectores en terreno)
       - `CENCO_Turnos` (Inicio y cierre de turno CENCO + enlace al documento PDF oficial guardado en Drive)
       - `CENCO_Alertas_CCTV` (Alertas visuales desde cámaras enviadas al Despachador con capturas y móvil sugerido)
       - `CENCO_Catalogo_Camaras` (Catálogo de cámaras CCTV por cuadrante, precargado con 15 cámaras base de Peñaflor)
       - `CENCO_Despachos` (Registro detallado de unidades despachadas por cuadrante y tiempos de respuesta)
   - Guarda los documentos **PDF de Cierre de Turno CENCO** en la carpeta `Peña_Seguro_1410_Actas_CENCO_PDF` de Google Drive.
   - Responde peticiones `GET` para sincronizar en tiempo real entre Inspectores, Administración y las consolas de **Telefonista 1410**, **Despachador** y **CCTV**.

---

## 🎧 Perfil Central CENCO CAD 1410 — Roles en Cadena

Todos los módulos existentes de **Inspectores** y **Administración** siguen funcionando exactamente igual, y ahora trabajan en conjunto con **CENCO**: cada procedimiento ingresado por un Inspector en terreno llega automáticamente a la **Bitácora Unificada CENCO (`CENCO_Bitacora_CAD`)** con su evidencia fotográfica.

| Rol CENCO | Usuario por Defecto | Clave Inicial | Función en Cadena |
| :--- | :--- | :--- | :--- |
| **1. Telefonista 1410** | `cenco.telefonista` | `Tel#1410` | Recibe llamados 1410/emergencias, categoriza (Prioridad P1–P4), adjunta imágenes y envía directo a la cola del Despachador. |
| **2. Despachador** | `cenco.despachador` | `Desp#1410` | Recibe llamados del Telefonista + Alertas de CCTV, visualiza móviles activos por cuadrante, despacha unidades y gestiona el ciclo del procedimiento. |
| **3. Operador CCTV** | `cenco.cctv` | `Cctv#1410` | Monitorea cámaras por cuadrante, consulta el radar de móviles cercanos a cada cámara y envía alertas visuales con capturas al Despachador. |
| **Supervisor CENCO** | `cenco.supervisor` | `Cenco#1410` | Acceso total a las 3 consolas, Bitácora Unificada, Reportes CAD y Actas PDF de Turno (también accesible desde el perfil `ADMIN`). |

> **Nota de despliegue en Apps Script:** Tras pegar el nuevo [`google-apps-script/Code.gs`](./google-apps-script/Code.gs), ejecute `inicializarHojas` una vez para crear las 5 nuevas pestañas CENCO y realice una **Nueva implementación** (o actualice la implementación web existente) para activar los nuevos endpoints.

---

## 🚀 Paso a Paso: Configurar el Excel en Google Drive (5 minutos)

1. Abre tu **Google Drive** y crea una nueva **Hoja de cálculo de Google (Google Sheets)** llamada, por ejemplo: `BD_Peña_Seguro_1410_Inspectores`.
2. En el menú superior de la planilla, haz clic en **Extensiones** → **Apps Script**.
3. Borra cualquier código que aparezca en el editor y pega el contenido completo del archivo [`google-apps-script/Code.gs`](./google-apps-script/Code.gs).
4. Haz clic en el ícono de **Guardar** (💾) y luego selecciona la función `inicializarHojas` en la barra superior y presiona **Ejecutar**:
   - Google te pedirá **Revisar permisos** → Selecciona tu cuenta → *Configuración avanzada* → *Ir a Proyecto (no seguro)* → **Permitir**.
   - Esto creará automáticamente las 4 pestañas con sus encabezados azules institucionales y la carpeta `Peña_Seguro_1410_Evidencias` en tu Google Drive.
5. En la esquina superior derecha de Apps Script, haz clic en **Implementar** → **Nueva implementación**:
   - En "Selecciona el tipo" (⚙️), elige **Aplicación web**.
   - **Descripción**: `API Peña Seguro 1410 v1`
   - **Ejecutar como**: `Yo` (tu cuenta de Google).
   - **Quién tiene acceso**: `Cualquier persona` (*Anyone* — necesario para que los celulares de los inspectores puedan enviar reportes sin pedir inicio de sesión de Google).
6. Haz clic en **Implementar** y copia la **URL de la aplicación web** (termina en `/exec`).
7. Conecta la URL en la App:
   - **Opción A (Desde la pantalla de la App)**: Abre la app, toca el botón **Excel Drive / Configurar (⚙️)** arriba a la derecha y pega la URL `/exec`.
   - **Opción B (Fija para todos en GitHub Pages)**: Pega la URL dentro de la constante `DEFAULT_GOOGLE_SCRIPT_URL` en `index.html` (línea ~651) antes de subir a GitHub.

---

## 🌐 Paso a Paso: Subir a GitHub y activar GitHub Pages

### Método 1: Desde el navegador web (Sin instalar Git en tu PC)
1. Ingresa a [github.com/new](https://github.com/new) e inicia sesión.
2. En **Repository name**, escribe: `pena-seguro-inspectores`.
3. Elige **Public** y haz clic en **Create repository**.
4. En la pantalla siguiente, haz clic en el enlace **"uploading an existing file"**.
5. Arrastra todos los archivos de esta carpeta (`index.html`, `pe_a_seguro_inspectores.html`, la carpeta `google-apps-script`, `README.md`, `.gitignore`) y presiona **Commit changes**.
6. Para tener el enlace web para los celulares:
   - Ve a la pestaña **Settings** del repositorio → menú izquierdo **Pages**.
   - En **Branch**, selecciona `main` (carpeta `/ (root)`) y presiona **Save**.
   - En 1 minuto tu aplicación estará publicada en:
     `https://<tu-usuario>.github.io/pena-seguro-inspectores/`

### Método 2: Desde terminal (Si instalas Git para Windows)
```bash
git init
git add .
git commit -m "feat: plataforma inspectores Peña Seguro 1410 con integración Google Sheets/Drive"
git branch -M main
git remote add origin https://github.com/<TU_USUARIO>/pena-seguro-inspectores.git
git push -u origin main
```
