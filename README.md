# Sales Dashboard - Nexus Corp

Dashboard ejecutivo de inteligencia comercial construido con Google Apps Script, Google Sheets, AppSheet y frontend en HTML/CSS/JS.

Repositorio: https://github.com/andresfayalac-cyber/Sales_dashboard

## 1) Resumen del proyecto

Este proyecto entrega un dashboard web para monitorear ventas, rentabilidad, clientes, productos y estrategia producto-mercado en una sola aplicacion.

El objetivo es pasar de reportes descriptivos a decisiones accionables:

- Que vender mas
- Donde enfocar la estrategia comercial
- Donde corregir mix y rentabilidad

Las vistas principales son:

- Overview
- Performance Pulse
- Clients
- Product & Market Strategy

## 2) Problema de negocio

Antes del dashboard, la organizacion tenia un problema comun:

- Datos de ventas distribuidos en hojas sin una capa de decision.
- Dificultad para entender rentabilidad real por cliente, producto y ubicacion.
- Bajo tiempo de reaccion para decisiones comerciales.
- Ausencia de una trazabilidad versionada del codigo y del despliegue.

## 3) Solucion implementada

### 3.1 Captura operacional con AppSheet

Se construyo una app en AppSheet para captura operativa de transacciones comerciales.

- Usuarios registran operaciones desde movil/web.
- La captura alimenta Google Sheets como base central.
- Se estandarizan campos para analitica (fecha, producto, ubicacion, importe, etc).

### 3.2 Base de datos en Google Sheets

Google Sheets funciona como capa de almacenamiento y transformacion inicial.

Tablas/hojas principales usadas por el dashboard:

- `fact_sales_clean` (hecho transaccional principal)
- `agg_monthly` (agregados mensuales)
- `agg_entity` (agregados por cliente)
- `kpi_summary` (KPIs ejecutivos)

### 3.3 Backend y web app en Google Apps Script

Se implemento backend serverless con Apps Script para:

- Exponer datos al frontend
- Aplicar reglas de negocio
- Manejar cache
- Publicar como Web App mediante `doGet`

Funciones clave:

- `doGet()`
- `getOverviewData()`
- `getMonthlyData()`
- `getTimeSeriesData()`
- `getEntityData()`
- `getProductLocationStrategyData()`
- `refreshAllDashboardData()`

### 3.4 Frontend del dashboard

Frontend modular por vistas HTML + estilos + JS:

- `index.html` (layout general)
- `ViewOverview.html`
- `ViewMonthly.html`
- `ViewEntities.html`
- `ViewProductsLocations.html`
- `Styles.html`
- `Scripts.html`

## 4) Resultados obtenidos

Resultados funcionales:

- Vista integral de performance comercial y financiera.
- Enfoque de decision en Product & Market Strategy (no solo rankings).
- Filtro temporal y analisis cruzado producto x ubicacion.
- Insights automaticos para apoyar accion comercial.

Resultados tecnicos:

- Publicacion como Web App lista para uso.
- Integracion de trabajo con `clasp`.
- Respaldo y versionado en GitHub.
- Mejoras de performance con carga por vista (lazy load).

## 5) Arquitectura y flujo end-to-end

1. Captura de datos en AppSheet.
2. Escritura en Google Sheets.
3. Apps Script procesa/agrega y expone endpoints internos.
4. Frontend del dashboard consume datos con `google.script.run`.
5. Dashboard renderiza visuales y tablas por vista.
6. Cambios se sincronizan con `clasp` y se respaldan en GitHub.

## 6) Estructura del proyecto

```text
Sales_dashboard/
|- appsscript.json
|- index.html
|- Styles.html
|- Scripts.html
|- ViewOverview.html
|- ViewMonthly.html
|- ViewEntities.html
|- ViewProductsLocations.html
|- Código.js (backend server-side)
|- .clasp.json (local)
|- .gitignore
```

## 7) Tecnologias usadas

- Google Apps Script (V8)
- Google Sheets
- AppSheet
- HTML5
- CSS3
- JavaScript (vanilla)
- ApexCharts
- clasp (CLI de Apps Script)
- Git + GitHub

## 8) Requisitos previos

- Cuenta de Google con acceso a Apps Script y Sheets.
- Node.js 18+ (recomendado) y npm.
- Git.
- `clasp` instalado globalmente.

## 9) Como ejecutar el proyecto desde GitHub

### 9.1 Clonar repositorio

```bash
git clone https://github.com/andresfayalac-cyber/Sales_dashboard.git
cd Sales_dashboard
```

### 9.2 Instalar dependencias de entorno

Este proyecto no requiere dependencias npm de aplicacion. Solo herramientas CLI:

```bash
npm install -g @google/clasp
```

### 9.3 Autenticacion y conexion con Apps Script

```bash
clasp login
```

Configura `.clasp.json` con tu `scriptId` si vas a apuntar a otro proyecto.

Ejemplo:

```json
{
  "scriptId": "TU_SCRIPT_ID",
  "rootDir": ""
}
```

### 9.4 Sincronizar codigo

Subir cambios locales al proyecto GAS:

```bash
clasp push
```

Bajar cambios desde GAS al repo local:

```bash
clasp pull
```

Ver estado:

```bash
clasp status
```

### 9.5 Despliegue como Web App

Crear version:

```bash
clasp version "release: dashboard"
```

Desplegar web app:

```bash
clasp deploy --description "webapp production"
```

Ver deployments:

```bash
clasp deployments
```

Abrir en navegador:

```bash
clasp open --webapp
```

## 10) Como correr todo el workflow (operativo + desarrollo)

### Flujo operativo (negocio)

1. Usuario registra data en AppSheet.
2. Data cae en Sheets (`fact_sales_clean`).
3. Dashboard lee datos y muestra indicadores.
4. Equipo revisa insights y toma decisiones.

### Flujo tecnico (desarrollo)

1. `git pull` para actualizar rama.
2. Editar archivos locales.
3. Probar en entorno de Apps Script.
4. `clasp push` para actualizar script.
5. Validar Web App.
6. `git add/commit/push` para respaldo en GitHub.

## 11) Metodologia

Se uso una metodologia iterativa orientada a producto y valor:

- Descubrimiento del problema y KPIs
- Diseno de modelo de datos operacional (AppSheet + Sheets)
- Construccion incremental de vistas analiticas
- Validacion funcional y ajuste UX
- Optimizacion de performance por etapas
- Versionado continuo y trazabilidad (GitHub + clasp)

## 12) Conclusiones y recomendaciones

### Conclusiones

- La integracion AppSheet + Sheets + Apps Script permite entregar BI accionable con bajo costo.
- Un dashboard orientado a decisiones mejora la lectura comercial frente a reportes descriptivos.
- El versionado con GitHub y sincronizacion con clasp hace el proyecto sostenible.

### Recomendaciones

- Mantener entornos separados (dev/staging/prod) con distintos `scriptId`.
- Definir monitoreo de tiempos de carga por vista.
- Agregar pruebas de regresion visual para cambios de UI responsive.
- Implementar checklist de despliegue para evitar errores manuales.

## 13) Datos de contacto

- Autor: Andres Ayala
- Rol: Data Scientist
- GitHub: https://github.com/andresfayalac-cyber
- Repositorio: https://github.com/andresfayalac-cyber/Sales_dashboard

Si deseas colaborar o reportar incidencias, abre un Issue en el repositorio.

## 14) Referencias

- Google Apps Script: https://developers.google.com/apps-script
- Web Apps en Apps Script: https://developers.google.com/apps-script/guides/web
- clasp: https://github.com/google/clasp
- Google AppSheet: https://about.appsheet.com
- ApexCharts: https://apexcharts.com/docs
- Google Sheets: https://support.google.com/docs
